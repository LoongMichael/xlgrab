"""
xlgrab核心Facade类，扩展pandas DataFrame功能
"""

import pandas as pd
import numpy as np
import re
from typing import Any, Optional, Union, List, Dict, Callable
import warnings

# 尝试导入openpyxl，如果失败则在使用时提示
try:
    from openpyxl.utils import coordinate_to_tuple
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


# 保存原始的DataFrame类
_OriginalDataFrame = pd.DataFrame
_OriginalSeries = pd.Series

class XlDataFrame(_OriginalDataFrame):
    """
    pandas DataFrame的增强版本，提供额外的便捷方法
    使用Facade模式，让DataFrame可以直接调用自定义方法
    """
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # 确保所有列都是XlSeries类型
        self._ensure_xl_columns()
    
    def _ensure_xl_columns(self):
        """确保所有列都是XlSeries类型"""
        for col in self.columns:
            if isinstance(self[col], _OriginalSeries) and not isinstance(self[col], XlSeries):
                self[col] = XlSeries(self[col])
    
    def __getitem__(self, key):
        """重写__getitem__方法，确保返回XlSeries"""
        result = super().__getitem__(key)
        if isinstance(result, _OriginalSeries) and not isinstance(result, XlSeries):
            return XlSeries(result)
        return result
    
    # ==================== Excel 范围操作 ====================
    
    def excel_range(self, *args, **kwargs):
        """Excel范围读取"""
        from .excel.range import excel_range
        return excel_range(self, *args, **kwargs)
    
    def offset_range(self, *args, **kwargs):
        """偏移范围选择"""
        from .excel.range import offset_range
        return offset_range(self, *args, **kwargs)
    
    def select_range(self, *args, **kwargs):
        """DSL风格区间选择"""
        from .excel.range import select_range
        return select_range(self, *args, **kwargs)
    
    # ==================== 数据操作 ====================
    
    def find_idx(self, *args, **kwargs):
        """数据查找"""
        from .data.search import find_idx_dataframe
        return find_idx_dataframe(self, *args, **kwargs)
    
    def apply_header(self, *args, **kwargs):
        """表头处理"""
        from .data.header import apply_header
        return apply_header(self, *args, **kwargs)
    
    # ==================== 数据探索方法 ====================
    # TODO: 在这里添加数据探索方法
    
    # ==================== 数据清洗方法 ====================
    # TODO: 在这里添加数据清洗方法
    
    # ==================== 数据转换方法 ====================
    # TODO: 在这里添加数据转换方法
    
    # ==================== 数据筛选方法 ====================
    # TODO: 在这里添加数据筛选方法
    
    # ==================== 数据聚合方法 ====================
    # TODO: 在这里添加数据聚合方法
    
    # ==================== 数据导出方法 ====================
    # TODO: 在这里添加数据导出方法


class XlSeries(_OriginalSeries):
    """pandas Series的增强版本"""
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
    
    # ==================== 数据操作 ====================
    
    def find_idx(self, *args, **kwargs):
        """数据查找"""
        from .data.search import find_idx_series
        return find_idx_series(self, *args, **kwargs)
    
    # ==================== Series 扩展方法 ====================
    # TODO: 在这里添加Series扩展方法