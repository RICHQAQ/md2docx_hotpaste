"""Main paste controller - orchestrates the entire conversion and insertion process."""

import traceback
import io

from md2docx_hotpaste.infra.process import detect_active_target

from ...services.clipboard import get_clipboard_text, is_clipboard_empty
from ...services.latex import convert_latex_delimiters
from ...services.pandoc import PandocConverter
from ...services.inserter.selector import TargetSelector
from ...services.inserter.word import WordInserter
from ...services.inserter.wps import WPSInserter
from ...services.inserter.excel import parse_markdown_table, MSExcelInserter, WPSExcelInserter
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
        self.ms_excel_inserter = MSExcelInserter()
        self.wps_excel_inserter = WPSExcelInserter()
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
            
            # 2. 获取剪贴板内容和配置
            md_text = get_clipboard_text()
            config = app_state.config
            
            # 3. 检测当前活动应用
            target = detect_active_target()
            log(f"Detected active target: {target}")
            
            # 4. 根据目标应用选择处理流程
            if target in ("excel", "wps_excel") and config.get("enable_excel", True):
                # Excel/WPS表格流程：直接插入表格数据
                self._handle_excel_flow(md_text, target, config)
            elif target in ("word", "wps", "none"):
                # Word/WPS文字流程：转换为DOCX后插入
                self._handle_word_flow(md_text, target, config)
            else:
                # 未知目标
                log(f"Unknown or unsupported target: {target}")
                self.notification_service.notify(
                    "MD2DOCX HotPaste",
                    "未检测到支持的应用，请打开 Word/WPS/Excel。",
                    ok=False
                )
            
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
    
    def _handle_excel_flow(self, md_text: str, target: str, config: dict) -> None:
        """
        Excel/WPS表格流程：解析Markdown表格并直接插入
        
        Args:
            md_text: Markdown文本
            target: 目标应用 (excel 或 wps_excel)
            config: 配置字典
        """
        # 根据目标选择插入器
        if target == "wps_excel":
            inserter = self.wps_excel_inserter
            app_name = "WPS 表格"
        else:  # excel
            inserter = self.ms_excel_inserter
            app_name = "Excel"
        
        # 解析Markdown表格
        table_data = parse_markdown_table(md_text)
        
        if table_data is None:
            # 不是有效的Markdown表格
            self.notification_service.notify(
                "MD2Excel HotPaste",
                f"未检测到有效的 Markdown 表格。\n当前应用: {app_name}",
                ok=False
            )
            return
        
        # 尝试插入表格
        log(f"Detected Markdown table with {len(table_data)} rows, inserting to {app_name}")
        try:
            keep_format = config.get("excel_keep_format", True)
            success = inserter.insert(table_data, keep_format=keep_format)
            
            if success:
                self.notification_service.notify(
                    "MD2Excel HotPaste",
                    f"已插入 {len(table_data)} 行表格到 {app_name}。",
                    ok=True
                )
        except InsertError as e:
            log(f"{app_name} insert failed: {e}")
            self.notification_service.notify(
                "MD2Excel HotPaste",
                f"插入到 {app_name} 失败。\n{str(e)}",
                ok=False
            )
    
    def _handle_word_flow(self, md_text: str, target: str, config: dict) -> None:
        """
        Word/WPS文字流程：转换Markdown为DOCX并插入
        
        Args:
            md_text: Markdown文本
            target: 目标应用 (word, wps 或 none)
            config: 配置字典
        """
        # 1. 处理LaTeX公式
        md_text = convert_latex_delimiters(md_text)
        
        # 2. 生成输出路径
        output_path = generate_output_path(
            keep_file=config.get("keep_file", False),
            save_dir=config.get("save_dir", "")
        )
        
        # 3. 转换为DOCX
        self._ensure_pandoc_converter()
        self.pandoc_converter.convert_to_docx(
            md_text=md_text,
            output_path=output_path,
            reference_docx=config.get("reference_docx")
        )
        log(f"Converted Markdown to DOCX: {output_path}")
        
        # 4. 根据配置确定最终插入目标
        configured_target = config.get("insert_target", "auto")
        if configured_target != "auto":
            # 用户指定了目标，使用TargetSelector解析
            target = self.target_selector.resolve_target(configured_target)
            log(f"User configured target: {configured_target}, resolved to: {target}")
        
        # 5. 执行插入
        inserted = self._perform_word_insertion(output_path, target)
        
        # 6. 清理文件
        self.cleanup_manager.cleanup_if_needed(
            file_path=output_path,
            keep_file=config.get("keep_file", False),
            insert_success=inserted,
            target=target
        )
        
        # 7. 显示结果通知
        self._show_word_result(target, inserted)
    
    def _ensure_pandoc_converter(self) -> None:
        """确保 Pandoc 转换器已初始化"""
        if self.pandoc_converter is None:
            pandoc_path = app_state.config.get("pandoc_path", "pandoc")
            self.pandoc_converter = PandocConverter(pandoc_path)
    
    def _perform_word_insertion(self, docx_path: str, target: str) -> bool:
        """
        执行Word/WPS文档插入
        
        Args:
            docx_path: DOCX文件路径
            target: 目标应用
            
        Returns:
            True 如果插入成功
        """
        if target == "none":
            log("Target is 'none', skip insertion")
            return False
        elif target == "word":
            try:
                return self.word_inserter.insert(docx_path)
            except InsertError as e:
                log(f"Word insertion failed: {e}")
                return False
        elif target == "wps":
            try:
                return self.wps_inserter.insert(docx_path)
            except InsertError as e:
                log(f"WPS insertion failed: {e}")
                return False
        else:
            log(f"Unknown insert target: {target}")
            return False
    
    def _show_word_result(self, target: str, inserted: bool) -> None:
        """显示Word/WPS流程的结果通知"""
        if target == "none":
            self.notification_service.notify(
                "MD2DOCX HotPaste",
                "已生成 DOCX（仅生成，不插入）。",
                ok=True
            )
        elif inserted:
            app_name = "Word" if target == "word" else "WPS 文字"
            self.notification_service.notify(
                "MD2DOCX HotPaste",
                f"已插入到 {app_name}。",
                ok=True
            )
        else:
            app_name = "Word" if target == "word" else "WPS 文字"
            self.notification_service.notify(
                "MD2DOCX HotPaste",
                f"未能插入到 {app_name}，请确认软件已打开且有光标。",
                ok=False
            )
