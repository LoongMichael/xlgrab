"""
xlgrab Excel 操作模块

提供Excel文件读取、范围操作、合并单元格处理等功能
"""

from .merger import unmerge_excel, unmerge_sheet
from .reader import read_excel_range as read_excel
from .range import excel_range, offset_range, select_range

__all__ = [
    'unmerge_excel',
    'unmerge_sheet', 
    'read_excel',
    'excel_range',
    'offset_range',
    'select_range',
]
