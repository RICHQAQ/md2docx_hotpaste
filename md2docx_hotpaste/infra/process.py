"""Process detection utilities."""

import os
import psutil
import win32gui
import win32process
from .logging import log


def get_foreground_process_name() -> str:
    """
    获取当前前台进程的名称
    
    Returns:
        进程名称（小写），失败时返回空字符串
    """
    try:
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            return ""
        
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(pid)
        return os.path.basename(process.exe()).lower()
        
    except Exception as e:
        log(f"Failed to get foreground process: {e}")
        return ""


def detect_active_target() -> str:
    """
    检测当前活跃的插入目标应用
    
    Returns:
        "word", "wps" 或空字符串
    """
    process_name = get_foreground_process_name()
    
    if "winword" in process_name:
        return "word"
    elif "wps" in process_name:
        return "wps"
    else:
        return ""
