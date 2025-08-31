"""WPS document insertion service."""

import win32com.client

from ...infra.com import ensure_com
from ...infra.logging import log
from ...core.errors import InsertError


class WPSInserter:
    """WPS 文档插入器"""
    
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
        # WPS 可能有多个不同的 ProgID
        prog_ids = ["kwps.Application", "wps.Application"]
        
        for prog_id in prog_ids:
            try:
                app = self._get_wps_application(prog_id)
                if app and self._try_insert_to_app(app, docx_path, prog_id):
                    log(f"Successfully inserted into WPS ({prog_id}): {docx_path}")
                    return True
            except Exception as e:
                log(f"Failed to insert via {prog_id}: {e}")
                continue
        
        raise InsertError("Failed to insert into WPS with all available ProgIDs")
    
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
    
    def _try_insert_to_app(self, app, docx_path: str, prog_id: str) -> bool:
        """尝试向 WPS 应用插入文档"""
        try:
            # 获取 Selection 对象
            selection = getattr(getattr(app, "ActiveWindow", app), "Selection", None)
            if selection is None:
                log(f"{prog_id} has no Selection object")
                return False
            
            # 执行插入
            selection.InsertFile(docx_path)
            return True
            
        except Exception as e:
            log(f"Insert failed for {prog_id}: {e}")
            return False
