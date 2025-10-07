"""
xlgrab扩展方法模块，提供更多pandas增强功能
"""

import pandas as pd
import numpy as np
from typing import Any, Optional, Union, List, Dict, Callable
import warnings


def register_extensions():
    """注册扩展方法到pandas DataFrame和Series"""
    
    # 注册到DataFrame
    from .core import XlDataFrame
    pd.DataFrame.find_idx = XlDataFrame.find_idx
    pd.DataFrame.excel_range = XlDataFrame.excel_range
    pd.DataFrame.offset_range = XlDataFrame.offset_range
    pd.DataFrame.select_range = XlDataFrame.select_range
    
    # 注册到Series
    from .core import XlSeries
    pd.Series.find_idx = XlSeries.find_idx


# ==================== DataFrame 扩展方法 ====================
# TODO: 在这里实现DataFrame扩展方法


# ==================== Series 扩展方法 ====================
# TODO: 在这里实现Series扩展方法