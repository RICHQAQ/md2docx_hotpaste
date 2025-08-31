"""Notification service."""

import os

try:
    from plyer import notification
    _NOTIFICATION_AVAILABLE = True
except ImportError:
    _NOTIFICATION_AVAILABLE = False

from ..core.constants import NOTIFICATION_TIMEOUT
from ..config.paths import get_app_icon_path
from ..infra.logging import log


class NotificationService:
    """通知服务"""
    
    def __init__(self, app_name: str = "MD2DOCX HotPaste"):
        self.app_name = app_name
        self.icon_path = get_app_icon_path()
    
    def notify(self, title: str, message: str, ok: bool = True) -> None:
        """
        发送系统通知
        
        Args:
            title: 通知标题
            message: 通知内容
            ok: 是否为成功状态（用于日志记录）
        """
        log(f"Notify: {title} - {message}")
        
        if not _NOTIFICATION_AVAILABLE:
            return
        
        try:
            notification.notify(
                title=title,
                message=message,
                timeout=NOTIFICATION_TIMEOUT,
                app_icon=self.icon_path if os.path.exists(self.icon_path) else None
            )
        except Exception as e:
            log(f"Notification error: {e}")
    
    def is_available(self) -> bool:
        """检查通知功能是否可用"""
        return _NOTIFICATION_AVAILABLE
