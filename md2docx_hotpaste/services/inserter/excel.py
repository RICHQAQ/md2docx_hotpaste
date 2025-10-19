"""Excel table insertion service."""

import re
from typing import List, Optional

from .base import BaseTableInserter
from ...core.errors import InsertError
from ...infra.logging import log


class CellFormat:
    """单元格格式信息"""
    def __init__(self, text: str):
        self.text = text
        self.bold = False
        self.italic = False
        self.strikethrough = False
        self.is_code = False
        self.clean_text = text
    
    def parse(self) -> str:
        """解析 Markdown 格式并记录格式信息"""
        text = self.text
        
        # 检测删除线 ~~text~~
        if re.search(r'~~.+?~~', text):
            self.strikethrough = True
            text = re.sub(r'~~(.+?)~~', r'\1', text)
        
        # 检测粗体 **text** 或 __text__
        if re.search(r'\*\*.+?\*\*', text) or re.search(r'__.+?__', text):
            self.bold = True
            text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
            text = re.sub(r'__(.+?)__', r'\1', text)
        
        # 检测斜体 *text* 或 _text_ (需要在粗体之后检测)
        if re.search(r'\*.+?\*', text) or re.search(r'_.+?_', text):
            self.italic = True
            text = re.sub(r'\*(.+?)\*', r'\1', text)
            text = re.sub(r'_(.+?)_', r'\1', text)
        
        # 检测代码 `code`
        if re.search(r'`.+?`', text):
            self.is_code = True
            text = re.sub(r'`(.+?)`', r'\1', text)
        
        # 移除链接 [text](url) -> text
        text = re.sub(r'\[(.+?)\]\(.+?\)', r'\1', text)
        
        self.clean_text = text.strip()
        return self.clean_text


def parse_markdown_table(md_text: str) -> Optional[List[List[str]]]:
    """
    解析 Markdown 表格为二维数组
    
    Args:
        md_text: Markdown 文本内容
        
    Returns:
        二维数组，每个元素代表一行的单元格内容；如果不是表格则返回 None
    """
    lines = md_text.strip().split('\n')
    if len(lines) < 2:
        return None
    
    table_data = []
    separator_found = False
    
    for i, line in enumerate(lines):
        line = line.strip()
        
        # 跳过空行
        if not line:
            continue
            
        # 检查是否为表格行（以 | 开头或结尾）
        if not (line.startswith('|') or line.endswith('|') or '|' in line):
            # 如果已经找到分隔符，说明表格结束
            if separator_found:
                break
            # 否则不是表格
            return None
        
        # 检查是否为分隔符行（如 |---|---|）
        if re.match(r'^\s*\|?\s*[-:]+\s*(\|\s*[-:]+\s*)+\|?\s*$', line):
            separator_found = True
            continue
        
        # 解析单元格
        cells = [cell.strip() for cell in line.split('|')]
        
        # 移除首尾的空元素（如果行是 |a|b| 格式）
        if cells and cells[0] == '':
            cells = cells[1:]
        if cells and cells[-1] == '':
            cells = cells[:-1]
        
        if cells:
            table_data.append(cells)
    
    # 必须找到分隔符才认为是有效表格
    if not separator_found or not table_data:
        return None
    
    return table_data


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
                            cell.Value = clean_text
                            
                            # 应用格式
                            try:
                                if cell_format.bold:
                                    cell.Font.Bold = True
                                if cell_format.italic:
                                    cell.Font.Italic = True
                                if cell_format.strikethrough:
                                    cell.Font.Strikethrough = True
                                if cell_format.is_code:
                                    # 代码使用等宽字体和浅灰色背景
                                    cell.Font.Name = "Consolas"
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
        获取 Excel 应用程序实例
        
        Returns:
            Excel 应用程序对象
            
        Raises:
            Exception: 无法获取实例时
        """
        import win32com.client
        
        try:
            # 尝试连接现有实例
            excel = win32com.client.GetActiveObject(self.prog_id)
            log(f"Successfully connected to {self.prog_id}")
            return excel
        except Exception as e:
            log(f"Failed to connect to {self.prog_id}: {e}")
            raise Exception(f"No {self.app_name} application found")
    
    # 保持向后兼容的别名
    _get_excel_application = _get_application
    
    def _clean_markdown_formatting(self, text: str) -> str:
        """
        清除 Markdown 格式符号
        
        Args:
            text: 包含 Markdown 格式的文本
            
        Returns:
            清除格式后的纯文本
        """
        # 移除粗体 **text** 或 __text__
        text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
        text = re.sub(r'__(.+?)__', r'\1', text)
        
        # 移除斜体 *text* 或 _text_
        text = re.sub(r'\*(.+?)\*', r'\1', text)
        text = re.sub(r'_(.+?)_', r'\1', text)
        
        # 移除行内代码 `code`
        text = re.sub(r'`(.+?)`', r'\1', text)
        
        # 移除链接 [text](url) -> text
        text = re.sub(r'\[(.+?)\]\(.+?\)', r'\1', text)
        
        # 移除删除线 ~~text~~
        text = re.sub(r'~~(.+?)~~', r'\1', text)
        
        return text.strip()


class MSExcelInserter(BaseExcelInserter):
    """Microsoft Excel 插入器"""
    
    def __init__(self):
        super().__init__(prog_id="Excel.Application", app_name="Excel")


class WPSExcelInserter(BaseExcelInserter):
    """WPS 表格插入器"""
    
    def __init__(self):
        super().__init__(prog_id="ket.Application", app_name="WPS 表格")
