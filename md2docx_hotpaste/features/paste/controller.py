"""Main paste controller - orchestrates the entire conversion and insertion process."""

import traceback
import io

from ...services.clipboard import get_clipboard_text, is_clipboard_empty
from ...services.latex import convert_latex_delimiters
from ...services.pandoc import PandocConverter
from ...services.inserter.selector import TargetSelector
from ...services.inserter.word import WordInserter
from ...services.inserter.wps import WPSInserter
from ...services.inserter.excel import parse_markdown_table, ExcelInserter
from ...services.notify import NotificationService
from ...infra.fs import generate_output_path
from ...infra.logging import log
from ...core.state import app_state
from ...core.errors import ClipboardError, PandocError, InsertError
from .cleanup import FileCleanupManager


class PasteController:
    """转换并插入控制器 - 业务流程编排"""
    
    def __init__(self):
        self.target_selector = TargetSelector()
        self.word_inserter = WordInserter()
        self.wps_inserter = WPSInserter()
        self.excel_inserter = ExcelInserter()
        self.cleanup_manager = FileCleanupManager()
        self.notification_service = NotificationService()
        self.pandoc_converter = None  # 延迟初始化
    
    def execute(self) -> None:
        """执行完整的转换和插入流程"""
        try:
            # 1. 检查剪贴板
            if is_clipboard_empty():
                self.notification_service.notify(
                    "MD2DOCX HotPaste",
                    "剪贴板为空，未处理。",
                    ok=False
                )
                return
            
            # 2. 获取剪贴板内容并处理
            md_text = get_clipboard_text()
            config = app_state.config
            
            # 2.1. 智能识别：如果是 Markdown 表格，尝试粘贴到 Excel
            if config.get("enable_excel", True):
                table_data = parse_markdown_table(md_text)
                if table_data is not None:
                    log("Detected Markdown table, trying to paste to Excel")
                    try:
                        keep_format = config.get("excel_keep_format", True)
                        success = self.excel_inserter.insert(table_data, keep_format=keep_format)
                        if success:
                            self.notification_service.notify(
                                "MD2Excel HotPaste",
                                f"已插入 {len(table_data)} 行表格到 Excel。",
                                ok=True
                            )
                            return
                    except InsertError as e:
                        # Excel 插入失败，继续尝试 Word/WPS 流程
                        log(f"Excel insert failed, fallback to Word/WPS: {e}")
            
            # 2.2. 继续原有的 Word/WPS 流程
            md_text = convert_latex_delimiters(md_text)
            
            # 3. 生成输出路径
            output_path = generate_output_path(
                keep_file=config.get("keep_file", False),
                save_dir=config.get("save_dir", "")
            )
            
            # 4. 转换为 DOCX
            self._ensure_pandoc_converter()
            self.pandoc_converter.convert_to_docx(
                md_text=md_text,
                output_path=output_path,
                reference_docx=config.get("reference_docx")
            )
            
            # 5. 确定插入目标
            configured_target = config.get("insert_target", "auto")
            target = self.target_selector.resolve_target(configured_target)
            
            # 6. 执行插入
            inserted = self._perform_insertion(output_path, target)
            
            # 7. 清理文件
            self.cleanup_manager.cleanup_if_needed(
                file_path=output_path,
                keep_file=config.get("keep_file", False),
                insert_success=inserted,
                target=target
            )
            
            # 8. 显示结果通知
            self._show_result_notification(target, inserted)
            
        except ClipboardError as e:
            log(f"Clipboard error: {e}")
            self.notification_service.notify(
                "MD2DOCX HotPaste",
                "剪贴板读取失败。",
                ok=False
            )
        except PandocError as e:
            log(f"Pandoc error: {e}")
            self.notification_service.notify(
                "MD2DOCX HotPaste",
                "Markdown 转换失败，请检查格式。",
                ok=False
            )
        except Exception:
            # 记录详细错误
            error_details = io.StringIO()
            traceback.print_exc(file=error_details)
            log(error_details.getvalue())
            
            self.notification_service.notify(
                "MD2DOCX HotPaste",
                "转换失败，请查看日志。",
                ok=False
            )
    
    def _ensure_pandoc_converter(self) -> None:
        """确保 Pandoc 转换器已初始化"""
        if self.pandoc_converter is None:
            pandoc_path = app_state.config.get("pandoc_path", "pandoc")
            self.pandoc_converter = PandocConverter(pandoc_path)
    
    def _perform_insertion(self, docx_path: str, target: str) -> bool:
        """执行文档插入"""
        if target == "none":
            return False
        elif target == "word":
            try:
                return self.word_inserter.insert(docx_path)
            except InsertError:
                return False
        elif target == "wps":
            try:
                return self.wps_inserter.insert(docx_path)
            except InsertError:
                return False
        else:
            log(f"Unknown insert target: {target}")
            return False
    
    def _show_result_notification(self, target: str, inserted: bool) -> None:
        """显示操作结果通知"""
        if target == "none":
            self.notification_service.notify(
                "MD2DOCX HotPaste",
                "已生成 DOCX（仅生成，不插入）。",
                ok=True
            )
        elif inserted:
            self.notification_service.notify(
                "MD2DOCX HotPaste",
                f"已插入到 {target.upper()}。",
                ok=True
            )
        else:
            self.notification_service.notify(
                "MD2DOCX HotPaste",
                f"未能插入到 {target.upper()}，请确认软件已打开且有光标。",
                ok=False
            )
