"""
xlgrab - A pandas enhancement library with Facade pattern
"""

# 导入核心类
from .core import XlDataFrame, XlSeries, _OriginalDataFrame, _OriginalSeries

# 导入扩展方法注册函数
from .extensions import register_extensions

# 导入工具函数
from .utils import *

# 替换pandas的DataFrame和Series为我们的增强版本
import pandas as pd
pd.DataFrame = XlDataFrame
pd.Series = XlSeries

# 自动注册扩展方法到pandas DataFrame和Series
register_extensions()

# 版本信息
__version__ = "0.1.0"
__author__ = "Your Name"
__email__ = "your.email@example.com"

# 导出主要类和函数
__all__ = [
    'XlDataFrame', 
    'XlSeries', 
    'register_extensions',
    'create_sample_data',
    'detect_data_types',
    'memory_usage_analysis',
    'find_duplicates',
    'outlier_detection',
    'data_quality_report',
    'smart_encoding',
    'feature_importance_analysis',
    'time_series_decomposition',
    'cross_validation_split',
    'export_to_multiple_formats'
]