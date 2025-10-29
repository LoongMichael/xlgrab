"""
xlgrab Excel 操作模块

提供Excel文件读取、范围操作、合并单元格处理、数据写入等功能
"""

from .merger import unmerge_excel, unmerge_sheet
from .reader import read_excel_range as read_excel
from .range import excel_range, offset_range, select_range
from .writer import write_to_excel, write_range_to_excel, to_sheet_many

__all__ = [
    'unmerge_excel',
    'unmerge_sheet', 
    'read_excel',
    'excel_range',
    'offset_range',
    'select_range',
    'write_to_excel',
    'write_range_to_excel',
    'to_sheet_many',
]
