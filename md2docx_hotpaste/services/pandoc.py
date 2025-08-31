"""Pandoc conversion service."""

import os
import subprocess
import tempfile
from typing import Optional

from ..core.errors import PandocError
from ..infra.logging import log


class PandocConverter:
    """Pandoc 转换器"""
    
    def __init__(self, pandoc_path: str = "pandoc"):
        self.pandoc_path = pandoc_path
    
    def convert_to_docx(
        self,
        md_text: str,
        output_path: str,
        reference_docx: Optional[str] = None
    ) -> None:
        """
        将 Markdown 文本转换为 DOCX 文件
        
        Args:
            md_text: Markdown 文本内容
            output_path: 输出 DOCX 文件路径
            reference_docx: 可选的参考文档模板路径
            
        Raises:
            PandocError: 转换失败时
        """
        with tempfile.TemporaryDirectory(prefix="md2docx_") as temp_dir:
            # 创建临时 Markdown 文件
            md_file = os.path.join(temp_dir, "input.md")
            with open(md_file, "w", encoding="utf-8", newline="\n") as f:
                f.write(md_text)
            
            # 构建 Pandoc 命令
            cmd = [
                self.pandoc_path, md_file,
                "--from", "markdown+tex_math_dollars+raw_tex",
                "--to", "docx",
                "-o", output_path,
                "--highlight-style", "tango"
            ]
            
            if reference_docx:
                cmd.extend(["--reference-doc", reference_docx])
            
            try:
                # 在 Windows 上隐藏控制台窗口
                startupinfo = None
                creationflags = 0
                if os.name == "nt":
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    creationflags = subprocess.CREATE_NO_WINDOW
                
                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    shell=False,
                    startupinfo=startupinfo,
                    creationflags=creationflags
                )
                
                if result.returncode != 0:
                    error_msg = result.stderr.strip() or result.stdout or "Pandoc conversion failed"
                    log(f"Pandoc error: {error_msg}")
                    raise PandocError(error_msg)
                    
            except FileNotFoundError:
                raise PandocError(f"Pandoc executable not found: {self.pandoc_path}")
            except Exception as e:
                log(f"Pandoc conversion failed: {e}")
                raise PandocError(f"Conversion failed: {e}")
