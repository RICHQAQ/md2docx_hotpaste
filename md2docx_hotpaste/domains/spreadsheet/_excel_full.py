"""Excel table insertion service."""

import re
from typing import List, Optional

from .base import BaseTableInserter
from ...core.errors import InsertError
from ...infra.logging import log


class TextSegment:
    """文本片段,带有格式信息"""
    def __init__(self, text: str, bold: bool = False, italic: bool = False,
                 strikethrough: bool = False, is_code: bool = False,
                 hyperlink_url: Optional[str] = None):
        self.text = text
        self.bold = bold
        self.italic = italic
        self.strikethrough = strikethrough
        self.is_code = is_code
        self.hyperlink_url = hyperlink_url  # 如果非空,表示这是一个超链接


class CellFormat:
    """单元格格式信息"""
    def __init__(self, text: str):
        self.text = text
        self.is_code_block = False
        self.has_newline = False
        self.segments = []  # List[TextSegment]
        self.clean_text = text
    
    def parse(self) -> str:
        """解析 Markdown 格式并生成文本片段(字符级解析)"""
        text = self.text
        
        # 处理 HTML 标签和换行
        text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
        if '\n' in text:
            self.has_newline = True
        
        # 检查是否包含代码块标签
        if '<pre>' in text.lower() or '<code>' in text.lower():
            self.is_code_block = True
            # 提取代码块内容
            text = re.sub(r'<pre>(.*?)</pre>',
                          lambda m: re.sub(r'<br\s*/?>', '\n', m.group(1), flags=re.IGNORECASE),
                          text, flags=re.DOTALL | re.IGNORECASE)
            text = re.sub(r'<code>(.*?)</code>',
                          lambda m: re.sub(r'<br\s*/?>', '\n', m.group(1), flags=re.IGNORECASE),
                          text, flags=re.DOTALL | re.IGNORECASE)
            self.clean_text = text.strip()
            self.segments = [TextSegment(self.clean_text, is_code=True)]
            return self.clean_text
        
        # 解析为文本片段
        self.segments = self._parse_segments(text)
        self.clean_text = ''.join(seg.text for seg in self.segments)
        return self.clean_text
    
    def _parse_segments(self, text: str, bold: bool = False, italic: bool = False,
                        strikethrough: bool = False) -> List[TextSegment]:
        """
        解析文本为带格式的片段列表
        
        Args:
            text: 要解析的文本
            bold: 当前是否在粗体环境中
            italic: 当前是否在斜体环境中
            strikethrough: 当前是否在删除线环境中
        """
        segments = []
        i = 0
        current_text = []
        
        def flush_current():
            """将当前累积的文本保存为片段"""
            if current_text:
                text_str = ''.join(current_text)
                if text_str:
                    segments.append(TextSegment(text_str, bold, italic, strikethrough))
                current_text.clear()
        
        while i < len(text):
            # 处理转义字符
            if text[i] == '\\' and i + 1 < len(text):
                current_text.append(text[i + 1])
                i += 2
                continue
            
            # 检测行内代码 `...`
            if text[i] == '`':
                end = text.find('`', i + 1)
                if end != -1:
                    flush_current()
                    # 代码块内的文本不做任何解析
                    code_text = text[i + 1:end]
                    segments.append(TextSegment(code_text, is_code=True))
                    i = end + 1
                    continue
            
            # 检测删除线 ~~...~~
            if text[i:i + 2] == '~~' and not strikethrough:
                end = text.find('~~', i + 2)
                if end != -1:
                    flush_current()
                    inner = text[i + 2:end]
                    # 递归解析,传递删除线状态
                    segments.extend(self._parse_segments(inner, bold, italic, True))
                    i = end + 2
                    continue
            
            # 检测粗斜体 ***...*** (必须在 ** 之前检测)
            if text[i:i + 3] == '***' and not bold and not italic:
                end = text.find('***', i + 3)
                if end != -1:
                    flush_current()
                    inner = text[i + 3:end]
                    # 递归解析,同时传递粗体和斜体状态
                    segments.extend(self._parse_segments(inner, True, True, strikethrough))
                    i = end + 3
                    continue
            
            # 检测粗体 **...** 或 __...__
            if text[i:i + 2] == '**' and not bold:
                # 查找匹配的 **,需要跳过单个 *
                end = i + 2
                while end < len(text) - 1:
                    if text[end:end + 2] == '**':
                        flush_current()
                        inner = text[i + 2:end]
                        # 递归解析,传递粗体状态
                        segments.extend(self._parse_segments(inner, True, italic, strikethrough))
                        i = end + 2
                        break
                    end += 1
                else:
                    # 没找到配对的,当普通字符处理
                    current_text.append(text[i])
                    i += 1
                continue
            
            # 检测粗斜体 ___...___ (必须在 __ 之前检测)
            if text[i:i + 3] == '___' and not bold and not italic:
                end = text.find('___', i + 3)
                if end != -1:
                    flush_current()
                    inner = text[i + 3:end]
                    # 递归解析,同时传递粗体和斜体状态
                    segments.extend(self._parse_segments(inner, True, True, strikethrough))
                    i = end + 3
                    continue
            
            if text[i:i + 2] == '__' and not bold:
                # 查找匹配的 __,需要跳过单个 _
                end = i + 2
                while end < len(text) - 1:
                    if text[end:end + 2] == '__':
                        flush_current()
                        inner = text[i + 2:end]
                        segments.extend(self._parse_segments(inner, True, italic, strikethrough))
                        i = end + 2
                        break
                    end += 1
                else:
                    # 没找到配对的,当普通字符处理
                    current_text.append(text[i])
                    i += 1
                continue
            
            # 检测斜体 *...* 或 _..._ (移除 not italic 限制,允许在粗体内使用斜体)
            if text[i] == '*' and (i + 1 >= len(text) or text[i + 1] != '*'):
                end = i + 1
                # 查找匹配的结束星号
                while end < len(text):
                    if text[end] == '*' and (end + 1 >= len(text) or text[end + 1] != '*'):
                        flush_current()
                        inner = text[i + 1:end]
                        # 递归解析,传递斜体状态
                        segments.extend(self._parse_segments(inner, bold, True, strikethrough))
                        i = end + 1
                        break
                    end += 1
                else:
                    # 没找到配对的,当普通字符处理
                    current_text.append(text[i])
                    i += 1
                continue
            
            if text[i] == '_' and (i + 1 >= len(text) or text[i + 1] != '_'):
                end = i + 1
                while end < len(text):
                    if text[end] == '_' and (end + 1 >= len(text) or text[end + 1] != '_'):
                        flush_current()
                        inner = text[i + 1:end]
                        segments.extend(self._parse_segments(inner, bold, True, strikethrough))
                        i = end + 1
                        break
                    end += 1
                else:
                    # 没找到配对的,当普通字符处理
                    current_text.append(text[i])
                    i += 1
                continue
            
            # 检测链接 [text](url)
            if text[i] == '[':
                close_bracket = text.find(']', i + 1)
                if close_bracket != -1 and close_bracket + 1 < len(text) and text[close_bracket + 1] == '(':
                    close_paren = text.find(')', close_bracket + 2)
                    if close_paren != -1:
                        flush_current()
                        # 提取链接文本和URL
                        link_text = text[i + 1:close_bracket]
                        link_url = text[close_bracket + 2:close_paren]
                        # 解析链接文本的格式,并添加超链接URL
                        link_segments = self._parse_segments(link_text, bold, italic, strikethrough)
                        # 为每个片段添加超链接URL
                        for seg in link_segments:
                            seg.hyperlink_url = link_url
                        segments.extend(link_segments)
                        i = close_paren + 1
                        continue
            
            # 普通字符
            current_text.append(text[i])
            i += 1
        
        flush_current()
        return segments


def _split_table_cells(line: str) -> List[str]:
    """
    按 | 分割表格单元格,正确处理转义的竖线
    
    Args:
        line: 表格行文本
        
    Returns:
        单元格列表
    """
    cells = []
    current_cell = []
    i = 0
    
    while i < len(line):
        if i > 0 and line[i] == '|' and line[i - 1] == '\\':
            # 转义的竖线,替换前一个反斜杠并添加竖线
            current_cell[-1] = '|'
            i += 1
        elif line[i] == '|':
            # 未转义的竖线,分割单元格
            cells.append(''.join(current_cell).strip())
            current_cell = []
            i += 1
        else:
            current_cell.append(line[i])
            i += 1
    
    # 添加最后一个单元格
    if current_cell or cells:  # 如果有内容或已经有单元格
        cells.append(''.join(current_cell).strip())
    
    return cells


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
        
        # 使用新的分割方法解析单元格
        cells = _split_table_cells(line)
        
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
