"""WPS document insertion service."""

import win32com.client

from .base import BaseDocumentInserter
from ...infra.com import ensure_com
from ...infra.logging import log
from ...core.errors import InsertError


class WPSInserter(BaseDocumentInserter):
    """WPS 文档插入器"""
    
    def __init__(self):
        # WPS 可能有多个不同的 ProgID，这里用主要的一个
        super().__init__(prog_id="kwps.Application", app_name="WPS 文字")
        self.prog_ids = ["kwps.Application", "KWPS.Application"]
    
    @ensure_com
    def insert(self, docx_path: str) -> bool:
        """
        将 DOCX 文件插入到 WPS 当前光标位置
        
        Args:
            docx_path: DOCX 文件路径
            
        Returns:
            True 如果插入成功
            
        Raises:
            InsertError: 插入失败时
        """
        # 尝试所有可能的 ProgID
        for prog_id in self.prog_ids:
            try:
                app = self._get_wps_application(prog_id)
                if app and self._perform_insertion(app, docx_path):
                    log(f"Successfully inserted into WPS ({prog_id}): {docx_path}")
                    return True
            except Exception as e:
                log(f"Failed to insert via {prog_id}: {e}")
                continue
        
        raise InsertError("Failed to insert into WPS with all available ProgIDs")
    
    def _get_application(self):
        """获取 WPS 应用程序实例（使用第一个可用的 ProgID）"""
        for prog_id in self.prog_ids:
            app = self._get_wps_application(prog_id)
            if app:
                return app
        raise Exception(f"No {self.app_name} application found")
    
    def _get_wps_application(self, prog_id: str):
        """获取 WPS 应用程序实例"""
        try:
            # 尝试连接现有实例
            return win32com.client.GetActiveObject(prog_id)
        except Exception:
            try:
                # 尝试创建新实例
                return win32com.client.Dispatch(prog_id)
            except Exception as e:
                log(f"Cannot get WPS application via {prog_id}: {e}")
                return None
    
    def _perform_insertion(self, app, docx_path: str) -> bool:
        """执行实际的插入操作"""
        try:
            # 获取 Selection 对象
            selection = getattr(getattr(app, "ActiveWindow", app), "Selection", None)
            if selection is None:
                log("WPS has no Selection object")
                return False
            
            # 执行插入
            selection.InsertFile(docx_path)
            return True
            
        except Exception as e:
            log(f"Insert failed for WPS: {e}")
            return False
