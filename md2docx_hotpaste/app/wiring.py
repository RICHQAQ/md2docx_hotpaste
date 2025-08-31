"""Dependency injection and object wiring."""

from ..config.loader import ConfigLoader
from ..services.notify import NotificationService
from ..features.paste.controller import PasteController
from ..ui.tray.menu import TrayMenuManager
from ..ui.tray.run import TrayRunner
from ..ui.hotkey.run import HotkeyRunner


class Container:
    """依赖注入容器"""
    
    def __init__(self):
        # 基础服务
        self.config_loader = ConfigLoader()
        self.notification_service = NotificationService()
        
        # 业务控制器
        self.paste_controller = PasteController()
        
        # UI 组件
        self.tray_menu_manager = TrayMenuManager(
            self.config_loader,
            self.notification_service
        )
        self.tray_runner = TrayRunner(self.tray_menu_manager)
        self.hotkey_runner = HotkeyRunner(self.paste_controller.execute)
        
        # 设置热键重启回调
        self.tray_menu_manager.set_restart_hotkey_callback(
            self.hotkey_runner.restart
        )
    
    def get_paste_controller(self) -> PasteController:
        return self.paste_controller
    
    def get_hotkey_runner(self) -> HotkeyRunner:
        return self.hotkey_runner
    
    def get_tray_runner(self) -> TrayRunner:
        return self.tray_runner
    
    def get_notification_service(self) -> NotificationService:
        return self.notification_service
