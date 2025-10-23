"""Windows window and process API utilities."""

import os
import psutil
import win32gui
import win32process
from ..logging import log


def get_foreground_window() -> int:
    """
    获取前台窗口句柄
    
    Returns:
        窗口句柄，失败时返回 0
    """
    try:
        return win32gui.GetForegroundWindow()
    except Exception as e:
        log(f"Failed to get foreground window: {e}")
        return 0


def get_foreground_process_name() -> str:
    """
    获取当前前台进程的名称
    
    Returns:
        进程名称（小写），失败时返回空字符串
    """
    try:
        hwnd = get_foreground_window()
        if not hwnd:
            return ""
        
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(pid)
        return os.path.basename(process.exe()).lower()
        
    except Exception as e:
        log(f"Failed to get foreground process: {e}")
        return ""


def get_foreground_window_title() -> str:
    """
    获取当前前台窗口标题
    
    Returns:
        窗口标题，失败时返回空字符串
    """
    try:
        hwnd = get_foreground_window()
        if not hwnd:
            return ""
        return win32gui.GetWindowText(hwnd)
    except Exception as e:
        log(f"Failed to get window title: {e}")
        return ""
