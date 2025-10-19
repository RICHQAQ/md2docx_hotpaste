"""WPS document insertion service."""

import win32com.client

from .word import BaseWordInserter
from ...infra.logging import log


class WPSInserter(BaseWordInserter):
    """WPS 文档插入器"""
    
    def __init__(self):
        # WPS 可能有多个不同的 ProgID
        super().__init__(
            prog_id=["kwps.Application", "KWPS.Application"],
            app_name="WPS 文字"
        )
    
    def _get_application(self):
        """获取 WPS 应用程序实例（尝试所有可能的 ProgID）"""
        for prog_id in self.prog_ids:
            try:
                # 尝试连接现有实例
                app = win32com.client.GetActiveObject(prog_id)
                log(f"Successfully connected to WPS via {prog_id}")
                return app
            except Exception:
                try:
                    # 尝试创建新实例
                    app = win32com.client.Dispatch(prog_id)
                    log(f"Successfully created WPS instance via {prog_id}")
                    return app
                except Exception as e:
                    log(f"Cannot get WPS application via {prog_id}: {e}")
                    continue
        
        raise Exception(f"未找到运行中的 {self.app_name}，请先打开")
    
    def _get_selection(self, app):
        """
        获取 WPS 的选择对象（兼容不同版本）
        
        Args:
            app: WPS 应用程序对象
            
        Returns:
            Selection 对象
        """
        # WPS 可能通过 ActiveWindow 访问 Selection
        try:
            active_window = getattr(app, "ActiveWindow", None)
            if active_window:
                selection = getattr(active_window, "Selection", None)
                if selection:
                    return selection
        except Exception:
            pass
        
        # 直接尝试从 app 获取 Selection
        return getattr(app, "Selection", None)
