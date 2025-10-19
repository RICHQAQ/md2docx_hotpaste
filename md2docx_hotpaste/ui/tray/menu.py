"""Tray menu construction and callbacks."""

import os
import pystray

from ...core.state import app_state
from ...config.loader import ConfigLoader
from ...config.paths import get_log_path, get_config_path
from ...services.notify import NotificationService
from ...infra.fs import ensure_dir
from ...infra.logging import log
from .icon import create_status_icon
from ..hotkey.dialog import HotkeyDialog


class TrayMenuManager:
    """托盘菜单管理器"""
    
    def __init__(self, config_loader: ConfigLoader, notification_service: NotificationService):
        self.config_loader = config_loader
        self.notification_service = notification_service
        self.restart_hotkey_callback = None  # 将由外部设置
    
    def set_restart_hotkey_callback(self, callback):
        """设置重启热键的回调函数"""
        self.restart_hotkey_callback = callback
    
    def build_menu(self) -> pystray.Menu:
        """构建托盘菜单"""
        config = app_state.config
        
        return pystray.Menu(
            # 快捷显示
            pystray.MenuItem(
                f"快捷键: {app_state.config['hotkey']}",
                lambda icon, item: None,
                enabled=False
            ),
            pystray.MenuItem(
                "启用热键",
                self._on_toggle_enabled,
                checked=lambda item: app_state.enabled
            ),
            pystray.MenuItem(
                "弹窗通知",
                self._on_toggle_notify,
                checked=lambda item: config.get("notify", True)
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("设置热键", self._on_set_hotkey),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                "插入文档目标",
                pystray.Menu(
                    pystray.MenuItem(
                        "Auto",
                        self._on_target_auto,
                        checked=lambda i: config.get("insert_target") == "auto"
                    ),
                    pystray.MenuItem(
                        "Word",
                        self._on_target_word,
                        checked=lambda i: config.get("insert_target") == "word"
                    ),
                    pystray.MenuItem(
                        "WPS",
                        self._on_target_wps,
                        checked=lambda i: config.get("insert_target") == "wps"
                    ),
                    pystray.MenuItem(
                        "None (仅生成)",
                        self._on_target_none,
                        checked=lambda i: config.get("insert_target") == "none"
                    ),
                )
            ),
            pystray.MenuItem(
                "保留生成文件",
                self._on_toggle_keep,
                checked=lambda item: config.get("keep_file", False)
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                "启动插入excel",
                self._on_toggle_excel,
                checked=lambda item: config.get("enable_excel", True)
            ),
            pystray.MenuItem(
                "启动excel解析特殊格式",
                self._on_toggle_excel_format,
                checked=lambda item: config.get("excel_keep_format", True)
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("打开保存目录", self._on_open_save_dir),
            pystray.MenuItem("查看日志", self._on_open_log),
            pystray.MenuItem("编辑配置", self._on_edit_config),
            pystray.MenuItem("重载配置/热键", self._on_reload),
            pystray.MenuItem("退出", self._on_quit)
        )
    
    # 菜单回调函数
    def _on_toggle_enabled(self, icon, item):
        """切换热键启用状态"""
        app_state.enabled = not app_state.enabled
        icon.icon = create_status_icon(ok=app_state.enabled)
        
        status = "已启用热键" if app_state.enabled else "已暂停热键"
        icon.menu = self.build_menu()
        self.notification_service.notify("MD2DOCX HotPaste", status, ok=app_state.enabled)
    
    def _on_set_hotkey(self, icon, item):
        """设置热键"""
        def save_hotkey(new_hotkey: str):
            """保存新热键并重启热键绑定"""
            try:
                # 更新配置
                app_state.config["hotkey"] = new_hotkey
                app_state.hotkey_str = new_hotkey
                self._save_config()
                
                # 重启热键绑定
                if self.restart_hotkey_callback:
                    self.restart_hotkey_callback()
                
                # 刷新菜单
                icon.menu = self.build_menu()
                
                log(f"Hotkey changed to: {new_hotkey}")
                self.notification_service.notify(
                    "MD2DOCX HotPaste",
                    f"热键已更新为：{new_hotkey}",
                    ok=True)
            except Exception as e:
                log(f"Failed to save hotkey: {e}")
                self.notification_service.notify(
                    "MD2DOCX HotPaste",
                    f"保存热键失败：{str(e)}",
                    ok=False)
                raise
        
        # 直接在主线程中显示对话框
        # tkinter 必须在主线程中运行，不能使用后台线程
        try:
            dialog = HotkeyDialog(
                current_hotkey=app_state.hotkey_str,
                on_save=save_hotkey
            )
            dialog.show()
        except Exception as e:
            log(f"Failed to show hotkey dialog: {e}")
            self.notification_service.notify("MD2DOCX HotPaste", f"打开热键设置失败：{str(e)}", ok=False)
    
    def _on_toggle_notify(self, icon, item):
        """切换通知状态"""
        current = app_state.config.get("notify", True)
        app_state.config["notify"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        if app_state.config["notify"]:
            self.notification_service.notify("MD2DOCX HotPaste", "已开启通知", ok=True)
        else:
            log("Notifications disabled via tray toggle")
    
    def _on_target_auto(self, icon, item):
        """设置插入目标为自动"""
        app_state.config["insert_target"] = "auto"
        self._save_config()
        icon.menu = self.build_menu()
        self.notification_service.notify("MD2DOCX HotPaste", "插入目标：Auto", ok=True)
    
    def _on_target_word(self, icon, item):
        """设置插入目标为 Word"""
        app_state.config["insert_target"] = "word"
        self._save_config()
        icon.menu = self.build_menu()
        self.notification_service.notify("MD2DOCX HotPaste", "插入目标：Word", ok=True)
    
    def _on_target_wps(self, icon, item):
        """设置插入目标为 WPS"""
        app_state.config["insert_target"] = "wps"
        self._save_config()
        icon.menu = self.build_menu()
        self.notification_service.notify("MD2DOCX HotPaste", "插入目标：WPS", ok=True)
    
    def _on_target_none(self, icon, item):
        """设置插入目标为无（仅生成）"""
        app_state.config["insert_target"] = "none"
        self._save_config()
        icon.menu = self.build_menu()
        self.notification_service.notify("MD2DOCX HotPaste", "仅生成，不插入", ok=True)
        
    def _on_toggle_excel(self, icon, item):
        """切换启用 Excel 插入"""
        current = app_state.config.get("enable_excel", True)
        app_state.config["enable_excel"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        self.notification_service.notify("MD2DOCX HotPaste", f"Excel 插入功能：{'开启' if not current else '关闭'}", ok=True)
        
    def _on_toggle_excel_format(self, icon, item):
        """切换 Excel 粘贴时是否保留格式"""
        current = app_state.config.get("excel_keep_format", True)
        app_state.config["excel_keep_format"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        self.notification_service.notify("MD2DOCX HotPaste", f"Excel 格式保留：{'开启' if not current else '关闭'}", ok=True)
    
    def _on_toggle_keep(self, icon, item):
        """切换保留文件状态"""
        current = app_state.config.get("keep_file", False)
        app_state.config["keep_file"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        status = "保留文件：开启" if app_state.config["keep_file"] else "保留文件：关闭"
        self.notification_service.notify("MD2DOCX HotPaste", status, ok=True)
    
    def _on_open_save_dir(self, icon, item):
        """打开保存目录"""
        save_dir = app_state.config.get("save_dir", "")
        save_dir = os.path.expandvars(save_dir)
        ensure_dir(save_dir)
        os.startfile(save_dir)
    
    def _on_open_log(self, icon, item):
        """打开日志文件"""
        log_path = get_log_path()
        if not os.path.exists(log_path):
            # 创建空日志文件
            open(log_path, "w", encoding="utf-8").close()
        os.startfile(log_path)
    
    def _on_edit_config(self, icon, item):
        """编辑配置文件"""
        config_path = get_config_path()
        if not os.path.exists(config_path):
            self._save_config()  # 创建默认配置文件
        os.startfile(config_path)
    
    def _on_reload(self, icon, item):
        """重载配置和热键"""
        try:
            app_state.config = self.config_loader.load()
            app_state.hotkey_str = app_state.config.get("hotkey", "<ctrl>+b")
            
            if self.restart_hotkey_callback:
                self.restart_hotkey_callback()
            icon.menu = self.build_menu()
            self.notification_service.notify("MD2DOCX HotPaste", "配置已重载", ok=True)
        except Exception as e:
            log(f"Failed to reload config: {e}")
            self.notification_service.notify("MD2DOCX HotPaste", "配置重载失败", ok=False)
    
    def _on_quit(self, icon, item):
        """退出应用程序"""
        icon.stop()
    
    def _save_config(self):
        """保存配置"""
        try:
            self.config_loader.save(app_state.config)
        except Exception as e:
            log(f"Failed to save config: {e}")
