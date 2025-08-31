"""Custom exceptions for the application."""


class MD2DOCXError(Exception):
    """应用程序基础异常"""
    pass


class ConfigError(MD2DOCXError):
    """配置相关异常"""
    pass


class PandocError(MD2DOCXError):
    """Pandoc 转换异常"""
    pass


class InsertError(MD2DOCXError):
    """文档插入异常"""
    pass


class ClipboardError(MD2DOCXError):
    """剪贴板操作异常"""
    pass
