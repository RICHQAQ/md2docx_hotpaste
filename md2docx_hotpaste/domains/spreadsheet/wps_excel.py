from .excel import BaseExcelInserter


class WPSExcelInserter(BaseExcelInserter):
    """WPS 表格插入器"""
    
    def __init__(self):
        super().__init__(prog_id=["ket.Application", "et.Application"], app_name="WPS 表格")