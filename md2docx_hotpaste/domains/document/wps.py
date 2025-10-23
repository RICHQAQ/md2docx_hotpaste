"""WPS document insertion."""

import time
import win32com.client

from .word import BaseWordInserter
from ...utils.logging import log
from ...core.constants import WORD_INSERT_RETRY_COUNT, WORD_INSERT_RETRY_DELAY


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
        WPS 通过文件关联打开时可能需要等待初始化，所以增加重试机制
        
        Args:
            app: WPS 应用程序对象
            
        Returns:
            Selection 对象
        """
        import pywintypes
        
        # 使用与 Word 插入相同的重试参数
        for attempt in range(WORD_INSERT_RETRY_COUNT):
            if attempt > 0:
                log(f"WPS Selection 获取失败，第 {attempt + 1} 次重试...")
                time.sleep(WORD_INSERT_RETRY_DELAY)
            
            # 方法1：直接从 app 获取 Selection（最常见的方式）
            try:
                selection = app.Selection
                if selection is not None:
                    log(f"获取 WPS Selection 成功（通过 app.Selection，尝试 {attempt + 1} 次）")
                    return selection
            except (AttributeError, pywintypes.com_error) as e:
                log(f"无法从 app 获取 Selection: {e}")
            
            # 方法2：通过 ActiveDocument.ActiveWindow.Selection
            try:
                selection = app.ActiveDocument.ActiveWindow.Selection
                if selection is not None:
                    log(f"获取 WPS Selection 成功（通过 ActiveDocument.ActiveWindow.Selection，尝试 {attempt + 1} 次）")
                    return selection
            except (AttributeError, pywintypes.com_error) as e:
                log(f"无法从 ActiveDocument.ActiveWindow 获取 Selection: {e}")
            
            # 方法3：通过 ActiveWindow.Selection
            try:
                selection = app.ActiveWindow.Selection
                if selection is not None:
                    log(f"获取 WPS Selection 成功（通过 ActiveWindow.Selection，尝试 {attempt + 1} 次）")
                    return selection
            except (AttributeError, pywintypes.com_error) as e:
                log(f"无法从 ActiveWindow 获取 Selection: {e}")
            
            # 方法4：通过 Documents(1).ActiveWindow.Selection
            try:
                documents = app.Documents
                if documents and documents.Count > 0:
                    selection = documents(1).ActiveWindow.Selection
                    if selection is not None:
                        log(f"获取 WPS Selection 成功（通过 Documents(1).ActiveWindow.Selection，尝试 {attempt + 1} 次）")
                        return selection
            except (AttributeError, pywintypes.com_error) as e:
                log(f"无法从 Documents(1).ActiveWindow 获取 Selection: {e}")
        
        log(f"所有获取 Selection 的方法都失败（已重试 {WORD_INSERT_RETRY_COUNT} 次）")
        return None
