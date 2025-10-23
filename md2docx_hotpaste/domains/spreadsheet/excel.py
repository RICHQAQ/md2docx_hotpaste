"""Excel and WPS spreadsheet inserters."""

from typing import List
from .base import BaseTableInserter
from .formatting import CellFormat
from ...core.errors import InsertError
from ...utils.logging import log


class BaseExcelInserter(BaseTableInserter):
    """Excel 表格插入器基类"""
    
    def insert(self, table_data: List[List[str]], keep_format: bool = True) -> bool:
        """
        将表格数据插入到 Excel 当前光标位置
        
        Args:
            table_data: 二维数组表格数据
            keep_format: 是否保留 Markdown 格式（粗体、斜体等）
            
        Returns:
            True 如果插入成功
            
        Raises:
            InsertError: 插入失败时
        """
        try:
            import pythoncom
            from pywintypes import com_error
            
            # 初始化 COM
            pythoncom.CoInitialize()
            
            try:
                # 获取 Excel 应用实例
                excel = self._get_excel_application()
            except Exception as e:
                raise InsertError(f"未找到运行中的 {self.app_name}，请先打开。错误: {e}")
            
            try:
                # 获取当前活动的工作表
                sheet = excel.ActiveSheet
                
                # 获取当前选中的单元格（起始位置）
                start_cell = excel.ActiveCell
                
                # 检查是否有活动单元格
                if start_cell is None:
                    raise InsertError(f"未选中任何单元格，请在 {self.app_name} 中点击要插入表格的起始位置")
                
                start_row = start_cell.Row
                start_col = start_cell.Column
                
                # 逐行逐列填充数据
                for i, row in enumerate(table_data):
                    for j, cell_value in enumerate(row):
                        cell = sheet.Cells(start_row + i, start_col + j)
                        
                        if keep_format:
                            # 解析格式并应用
                            cell_format = CellFormat(cell_value)
                            clean_text = cell_format.parse()
                            
                            # 应用格式
                            try:
                                # 如果包含换行,启用单元格自动换行
                                if cell_format.has_newline:
                                    cell.WrapText = True
                                
                                if cell_format.is_code_block:
                                    # 代码块使用等宽字体、浅灰色背景
                                    cell.Value = clean_text
                                    cell.Font.Name = "Consolas"
                                    cell.Interior.Color = 0xF0F0F0  # 浅灰色
                                    cell.WrapText = True
                                    # 设置垂直对齐为顶部
                                    cell.VerticalAlignment = -4160  # xlTop
                                elif len(cell_format.segments) == 0:
                                    # 没有格式,直接设置
                                    cell.Value = clean_text
                                elif len(cell_format.segments) == 1 and not any([
                                    cell_format.segments[0].bold,
                                    cell_format.segments[0].italic,
                                    cell_format.segments[0].strikethrough,
                                    cell_format.segments[0].is_code,
                                    cell_format.segments[0].hyperlink_url
                                ]):
                                    # 单个片段且无格式无链接,直接设置
                                    cell.Value = clean_text
                                else:
                                    # 检查是否整个单元格都是一个超链接
                                    if (len(cell_format.segments) == 1
                                            and cell_format.segments[0].hyperlink_url
                                            and not any([cell_format.segments[0].bold,
                                                        cell_format.segments[0].italic,
                                                        cell_format.segments[0].strikethrough])):
                                        # 单个超链接,使用 Hyperlinks.Add
                                        segment = cell_format.segments[0]
                                        cell.Value = segment.text
                                        try:
                                            sheet.Hyperlinks.Add(
                                                Anchor=cell,
                                                Address=segment.hyperlink_url,
                                                TextToDisplay=segment.text
                                            )
                                        except com_error as e:
                                            log(f"Failed to add hyperlink: {e}")
                                    else:
                                        # 使用富文本格式
                                        cell.Value = clean_text
                                        char_index = 1  # Excel 字符索引从1开始
                                        
                                        # 检查是否有超链接(有超链接时不能使用 GetCharacters)
                                        has_hyperlink = any(seg.hyperlink_url for seg in cell_format.segments)
                                        
                                        if has_hyperlink:
                                            # 有超链接时,只设置文本,不设置富文本格式
                                            # 因为 Hyperlinks.Add 和 GetCharacters 不兼容
                                            for segment in cell_format.segments:
                                                if segment.hyperlink_url:
                                                    # 为链接部分添加超链接
                                                    # 注意: Excel 单元格只能有一个超链接,这里取第一个
                                                    try:
                                                        sheet.Hyperlinks.Add(
                                                            Anchor=cell,
                                                            Address=segment.hyperlink_url,
                                                            TextToDisplay=clean_text
                                                        )
                                                        break
                                                    except com_error as e:
                                                        log(f"Failed to add hyperlink: {e}")
                                        else:
                                            # 没有超链接,可以使用富文本格式
                                            for segment in cell_format.segments:
                                                if not segment.text:
                                                    continue
                                                
                                                seg_len = len(segment.text)
                                                # 获取字符范围
                                                chars = cell.GetCharacters(char_index, seg_len)
                                                
                                                # 应用格式
                                                if segment.is_code:
                                                    chars.Font.Name = "Consolas"
                                                if segment.bold:
                                                    chars.Font.Bold = True
                                                if segment.italic:
                                                    chars.Font.Italic = True
                                                if segment.strikethrough:
                                                    chars.Font.Strikethrough = True
                                                
                                                char_index += seg_len
                                        
                                        # 如果有行内代码,设置整个单元格背景
                                        if any(seg.is_code for seg in cell_format.segments):
                                            cell.Interior.Color = 0xF0F0F0  # 浅灰色
                                        
                            except com_error as e:
                                # 格式应用失败，记录但继续
                                log(f"Failed to apply format to cell ({i},{j}): {e}")
                        else:
                            # 仅清除格式
                            cell.Value = self._clean_markdown_formatting(cell_value)
                
                # 选中插入的区域（可选）
                end_row = start_row + len(table_data) - 1
                end_col = start_col + max(len(row) for row in table_data) - 1
                range_to_select = sheet.Range(
                    sheet.Cells(start_row, start_col),
                    sheet.Cells(end_row, end_col)
                )
                range_to_select.Select()
                
                log(f"Successfully inserted table to {self.app_name}: {len(table_data)} rows, keep_format={keep_format}")
                return True
                
            finally:
                pythoncom.CoUninitialize()
                
        except InsertError:
            raise
        except Exception as e:
            log(f"Failed to insert table to {self.app_name}: {e}")
            raise InsertError(f"{self.app_name} 插入失败: {e}")
    
    def _get_application(self):
        """
        获取 Excel 应用程序实例（尝试所有可能的 ProgID）
        
        Returns:
            Excel 应用程序对象
            
        Raises:
            Exception: 无法获取实例时
        """
        import win32com.client
        
        # 尝试所有可能的 ProgID
        for prog_id in self.prog_ids:
            try:
                # 尝试连接现有实例
                excel = win32com.client.GetActiveObject(prog_id)
                log(f"Successfully connected to {prog_id}")
                return excel
            except Exception as e:
                log(f"Failed to connect to {prog_id}: {e}")
                continue
        
        raise Exception(f"未找到运行中的 {self.app_name}，请先打开")
    
    # 保持向后兼容的别名
    _get_excel_application = _get_application
    
    def _clean_markdown_formatting(self, text: str) -> str:
        """
        清除 Markdown 格式符号(使用字符级解析)
        
        Args:
            text: 包含 Markdown 格式的文本
            
        Returns:
            清除格式后的纯文本
        """
        # 直接使用 CellFormat 来解析
        cell_format = CellFormat(text)
        return cell_format.parse()


class MSExcelInserter(BaseExcelInserter):
    """Microsoft Excel 插入器"""
    
    def __init__(self):
        super().__init__(prog_id="Excel.Application", app_name="Excel")


class WPSExcelInserter(BaseExcelInserter):
    """WPS 表格插入器"""
    
    def __init__(self):
        super().__init__(prog_id="ket.Application", app_name="WPS 表格")
