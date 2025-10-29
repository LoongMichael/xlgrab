"""
xlgrab - A pandas enhancement library with Facade pattern
"""

# 导入核心类
from .core import XlDataFrame, XlSeries, _OriginalDataFrame, _OriginalSeries

# 导入 accessor 注册（替代直接替换类）
from .accessors import enable_direct_methods  # re-export optional helper

# 导入扩展注册
from .extensions import register_extensions

# 导入Excel功能
from .excel import (
    unmerge_excel, 
    unmerge_sheet, 
    read_excel, 
    write_to_excel, 
    write_range_to_excel, 
    to_sheet_many
)

# 默认不替换 pandas 类，改为通过 pandas Accessor 暴露功能：
# df.xl.excel_range / s.xl.find_idx 等。
# 如果用户想启用直呼 df.excel_range，可调用 xlgrab.enable_direct_methods()。

# 版本信息
__version__ = "0.1.0"
__author__ = "Your Name"
__email__ = "your.email@example.com"

# 导出主要类和函数
__all__ = [
    'XlDataFrame', 
    'XlSeries', 
    'register_extensions',
    'enable_direct_methods',
    'unmerge_excel',
    'unmerge_sheet',
    'read_excel',
    'write_to_excel',
    'write_range_to_excel',
    'to_sheet_many'
]
