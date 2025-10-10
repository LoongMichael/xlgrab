"""
xlgrab 数据处理模块

提供数据查找、表头处理、数据转换等功能
"""

from .search import find_idx_dataframe, find_idx_series
from .header import apply_header

__all__ = [
    'find_idx_dataframe',
    'find_idx_series',
    'apply_header',
]
