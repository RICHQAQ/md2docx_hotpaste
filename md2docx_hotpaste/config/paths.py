"""Resource and file path management."""

import os
import sys


def get_base_dir() -> str:
    """获取应用程序基础目录"""
    # 返回项目根目录（md2docx_hotpaste）
    current_file = os.path.abspath(__file__)
    # 从 md2docx_hotpaste/config/paths.py 回到 md2docx_hotpaste/
    return os.path.dirname(os.path.dirname(os.path.dirname(current_file)))


def resource_path(relative_path: str) -> str:
    """获取资源文件路径（支持 PyInstaller）"""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(get_base_dir(), relative_path)


def get_config_path() -> str:
    """获取配置文件路径"""
    return os.path.join(get_base_dir(), "config.json")


def get_log_path() -> str:
    """获取日志文件路径"""
    return os.path.join(get_base_dir(), "md2docx.log")


def get_app_icon_path() -> str:
    """获取应用图标路径 (.ico)"""
    return resource_path(os.path.join("assets", "icons", "logo.ico"))


def get_app_png_path() -> str:
    """获取应用图标路径 (.png)"""
    return resource_path(os.path.join("assets", "icons", "logo.png"))
