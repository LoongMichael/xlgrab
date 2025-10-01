"""
xlgrab - 极简Excel数据提取库

设计理念：
- 极简API：用户只需关心"提取什么数据"
- 函数式设计：纯函数，无状态，易测试
- 渐进式复杂度：从简单到复杂，按需使用

使用示例：
    # 简单提取
    result = xlgrab.extract_simple("file.xlsx", "Sheet1", "A1:C10")
    
    # 带表头提取
    result = xlgrab.extract_with_header("file.xlsx", "Sheet1", "A1:C1", "A2:C10")
    
    # 锚点提取
    result = xlgrab.extract("file.xlsx", [
        xlgrab.anchor_spec("Sheet1", "A", "姓名", 1, (1, 0))
    ])
"""

from .core import (
    # 核心函数
    extract,
    extract_simple,
    extract_with_header,
    extract_table,
    extract_list,
    
    # 区域定义
    range_spec,
    anchor_spec,
    
    # 结果类型
    ExtractResult,
)

# 保持向后兼容
from .reader import get_sheet, get_region, get_cell, last_data_row

__version__ = "1.0.0"
__all__ = [
    # 核心函数
    "extract",
    "extract_simple", 
    "extract_with_header",
    "extract_table",
    "extract_list",
    
    # 区域定义
    "range_spec",
    "anchor_spec",
    
    # 结果类型
    "ExtractResult",
    
    # 底层函数（向后兼容）
    "get_sheet",
    "get_region", 
    "get_cell",
    "last_data_row",
]