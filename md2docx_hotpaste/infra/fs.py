"""File system utilities."""

import os
import pathlib
import tempfile
from datetime import datetime


def ensure_dir(path: str) -> None:
    """确保目录存在，如不存在则创建"""
    pathlib.Path(path).mkdir(parents=True, exist_ok=True)


def generate_output_path(keep_file: bool, save_dir: str) -> str:
    """
    生成输出文件路径
    
    Args:
        keep_file: 是否保留文件
        save_dir: 保存目录
        
    Returns:
        输出文件的完整路径
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"md_paste_{timestamp}.docx"
    
    if keep_file:
        ensure_dir(save_dir)
        return os.path.join(save_dir, filename)
    else:
        return os.path.join(tempfile.gettempdir(), filename)
