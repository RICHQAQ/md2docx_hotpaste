"""Word document insertion service."""

import time
import win32com.client
from win32com.client import gencache

from .base import BaseDocumentInserter
from ...infra.com import ensure_com
from ...infra.logging import log
from ...core.constants import WORD_INSERT_RETRY_COUNT, WORD_INSERT_RETRY_DELAY
from ...core.errors import InsertError


class WordInserter(BaseDocumentInserter):
    """Word 文档插入器"""
    
    def __init__(self):
        super().__init__(prog_id="Word.Application", app_name="Word")
    
    @ensure_com
    def insert(self, docx_path: str) -> bool:
        """
        将 DOCX 文件插入到 Word 当前光标位置
        
        Args:
            docx_path: DOCX 文件路径
            
        Returns:
            True 如果插入成功
            
        Raises:
            InsertError: 插入失败时
        """
        try:
            app = self._get_application()
            return self._perform_insertion(app, docx_path)
        except Exception as e:
            log(f"Word insertion failed: {e}")
            raise InsertError(f"Word insertion failed: {e}")
    
    def _get_application(self):
        """获取 Word 应用程序实例"""
        try:
            # 尝试连接现有的 Word 实例
            app = win32com.client.GetActiveObject(self.prog_id)
        except Exception:
            app = gencache.EnsureDispatch(self.prog_id)
        
        # 确保 Word 可见并有文档
        self._ensure_word_ready(app)
        return app
    
    def _perform_insertion(self, app, docx_path: str) -> bool:
        """执行实际的插入操作"""
        # 获取当前选择区域
        selection = getattr(app, "Selection", None)
        if selection is None:
            raise InsertError("Cannot access Word selection")
        
        range_obj = selection.Range
        
        # 重试插入文件
        for attempt in range(WORD_INSERT_RETRY_COUNT):
            try:
                range_obj.InsertFile(docx_path)
                log(f"Successfully inserted into Word: {docx_path}")
                return True
            except Exception as e:
                if attempt < WORD_INSERT_RETRY_COUNT - 1:
                    log(f"Word insert attempt {attempt + 1} failed, retrying: {e}")
                    time.sleep(WORD_INSERT_RETRY_DELAY)
                else:
                    raise InsertError(f"Failed to insert after {WORD_INSERT_RETRY_COUNT} attempts: {e}")
        
        return False
    
    def _ensure_word_ready(self, app) -> None:
        """确保 Word 应用程序处于就绪状态"""
        try:
            # 确保 Word 可见
            app.Visible = True
        except Exception:
            pass
        
        # 确保有打开的文档
        documents = getattr(app, "Documents", None)
        if documents is None or documents.Count == 0:
            documents.Add()  # 创建新文档
        
        # 切换到主文档正文（避免停留在页眉/页脚/导航窗格）
        try:
            # 0 = wdSeekMainDocument
            app.ActiveWindow.View.SeekView = 0
        except Exception:
            pass
