"""Application entry point and initialization."""

try:
    import ctypes
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("MD2DOCX.HotPaste")
except Exception:
    pass

from ..core.state import app_state
from ..config.loader import ConfigLoader
from ..config.paths import get_app_icon_path
from ..infra.logging import log
from ..services.notify import NotificationService
from .wiring import Container


def initialize_application() -> Container:
    """初始化应用程序"""
    # 1. 加载配置
    config_loader = ConfigLoader()
    config = config_loader.load()
    app_state.config = config
    app_state.hotkey_str = config.get("hotkey", "<ctrl>+b")
    
    # 2. 创建依赖注入容器
    container = Container()
    
    log("Application initialized successfully")
    return container


def show_startup_notification(notification_service: NotificationService) -> None:
    """显示启动通知"""
    try:
        # 确保图标路径存在（仅用于验证）
        get_app_icon_path()
        notification_service.notify(
            "MD2DOCX HotPaste",
            "启动成功，已经运行在后台。",
            ok=True
        )
    except Exception as e:
        log(f"Failed to show startup notification: {e}")


def main() -> None:
    """应用程序主入口点"""
    try:
        # 初始化应用程序
        container = initialize_application()
        
        # 启动热键监听
        hotkey_runner = container.get_hotkey_runner()
        hotkey_runner.start()
        
        # 显示启动通知
        notification_service = container.get_notification_service()
        show_startup_notification(notification_service)
        
        # 启动托盘（阻塞运行）
        tray_runner = container.get_tray_runner()
        tray_runner.run()
        
    except KeyboardInterrupt:
        log("Application interrupted by user")
    except Exception as e:
        log(f"Fatal error: {e}")
        raise
    finally:
        log("Application shutting down")


if __name__ == "__main__":
    main()
