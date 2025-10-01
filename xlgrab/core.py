"""
xlgrab 核心模块 - 极简Excel数据提取

设计原则：
- 极简API：用户只需关心"提取什么数据"
- 函数式设计：纯函数，无状态，易测试
- 单一职责：每个函数只做一件事
- 渐进式复杂度：从简单到复杂，按需使用
"""

from typing import List, Dict, Any, Optional, Union
import pandas as pd
from .reader import get_sheet, get_region, get_cell, last_data_row


# ============================================================================
# 核心数据类型
# ============================================================================

class ExtractResult:
    """提取结果 - 极简设计"""
    def __init__(self, data: List[List[Any]], columns: List[str] = None, errors: List[str] = None):
        self.data = data
        self.columns = columns or []
        self.errors = errors or []
    
    def to_dataframe(self) -> pd.DataFrame:
        """转换为pandas DataFrame"""
        if not self.data:
            return pd.DataFrame()
        return pd.DataFrame(self.data, columns=self.columns)
    
    def to_dict(self) -> List[Dict[str, Any]]:
        """转换为字典列表"""
        if not self.data or not self.columns:
            return []
        return [dict(zip(self.columns, row)) for row in self.data]


# ============================================================================
# 区域定义 - 极简语法
# ============================================================================

def range_spec(sheet: str, area: str) -> Dict[str, str]:
    """定义区域 - 极简语法
    
    Args:
        sheet: 工作表名称
        area: 区域描述，支持多种格式：
            - "A1:C10" - 固定区域
            - "A1:last" - 从A1到最后一行
            - "A1:lastcol" - 从A1到最后一列
            - "A1:lastlast" - 从A1到最后一行最后一列
    
    Returns:
        区域规范字典
    """
    return {"sheet": sheet, "area": area}


def anchor_spec(sheet: str, column: str, text: str, occurrence: int = 1, offset: tuple = (0, 0)) -> Dict[str, Any]:
    """定义锚点区域 - 通过文本查找
    
    Args:
        sheet: 工作表名称
        column: 搜索列（如"A", "B"）
        text: 要查找的文本
        occurrence: 第几次出现（默认1）
        offset: 偏移量 (行偏移, 列偏移)
    
    Returns:
        锚点规范字典
    """
    return {
        "sheet": sheet,
        "type": "anchor",
        "column": column,
        "text": text,
        "occurrence": occurrence,
        "offset": offset
    }


# ============================================================================
# 核心提取函数 - 极简API
# ============================================================================

def extract(file_path: str, specs: List[Dict[str, Any]]) -> ExtractResult:
    """提取Excel数据 - 主入口函数
    
    Args:
        file_path: Excel文件路径
        specs: 提取规范列表，每个规范定义要提取的区域
    
    Returns:
        ExtractResult: 提取结果
    """
    all_data = []
    all_columns = []
    all_errors = []
    
    for spec in specs:
        try:
            result = _extract_single(file_path, spec)
            if result.data:
                all_data.extend(result.data)
            if result.columns and not all_columns:
                all_columns = result.columns
            all_errors.extend(result.errors)
        except Exception as e:
            all_errors.append(f"提取失败: {str(e)}")
    
    return ExtractResult(all_data, all_columns, all_errors)


def extract_simple(file_path: str, sheet: str, area: str) -> ExtractResult:
    """简单提取 - 最常用的函数
    
    Args:
        file_path: Excel文件路径
        sheet: 工作表名称
        area: 区域描述，如"A1:C10"
    
    Returns:
        ExtractResult: 提取结果
    """
    spec = range_spec(sheet, area)
    return extract(file_path, [spec])


def extract_with_header(file_path: str, sheet: str, header_area: str, data_area: str) -> ExtractResult:
    """带表头的提取 - 常用模式
    
    Args:
        file_path: Excel文件路径
        sheet: 工作表名称
        header_area: 表头区域，如"A1:C1"
        data_area: 数据区域，如"A2:C10"
    
    Returns:
        ExtractResult: 提取结果，包含列名
    """
    # 提取表头
    header_result = extract_simple(file_path, sheet, header_area)
    if not header_result.data:
        return ExtractResult([], [], ["表头提取失败"])
    
    # 提取数据
    data_result = extract_simple(file_path, sheet, data_area)
    if not data_result.data:
        return ExtractResult([], [], ["数据提取失败"])
    
    # 合并结果
    columns = [str(cell) for cell in header_result.data[0]]
    return ExtractResult(data_result.data, columns, header_result.errors + data_result.errors)


# ============================================================================
# 内部实现函数
# ============================================================================

def _extract_single(file_path: str, spec: Dict[str, Any]) -> ExtractResult:
    """提取单个区域"""
    sheet = spec["sheet"]
    
    if "area" in spec:
        return _extract_range(file_path, sheet, spec["area"])
    elif spec.get("type") == "anchor":
        return _extract_anchor(file_path, spec)
    else:
        return ExtractResult([], [], [f"未知的提取类型: {spec}"])


def _extract_range(file_path: str, sheet: str, area: str) -> ExtractResult:
    """提取固定区域"""
    try:
        # 解析区域
        start_cell, end_cell = area.split(":")
        
        # 处理特殊区域
        if end_cell == "last":
            end_row = last_data_row(file_path, sheet, 1)
            end_cell = f"{start_cell[0]}{end_row}"
        elif end_cell == "lastcol":
            # 简化处理，假设到Z列
            end_cell = f"Z{start_cell[1:]}"
        elif end_cell == "lastlast":
            end_row = last_data_row(file_path, sheet, 1)
            end_cell = f"Z{end_row}"
        
        # 转换为行列坐标
        start_row, start_col = _a1_to_rc(start_cell)
        end_row, end_col = _a1_to_rc(end_cell)
        
        # 提取数据
        df = get_region(file_path, sheet, start_row, end_row, start_col, end_col)
        if df is None:
            return ExtractResult([], [], [f"无法读取区域 {area}"])
        
        data = df.values.tolist()
        return ExtractResult(data)
        
    except Exception as e:
        return ExtractResult([], [], [f"区域解析失败: {str(e)}"])


def _extract_anchor(file_path: str, spec: Dict[str, Any]) -> ExtractResult:
    """提取锚点区域"""
    try:
        sheet = spec["sheet"]
        column = spec["column"]
        text = spec["text"]
        occurrence = spec.get("occurrence", 1)
        offset = spec.get("offset", (0, 0))
        
        # 获取整个sheet
        df = get_sheet(file_path, sheet)
        if df is None:
            return ExtractResult([], [], [f"无法读取工作表 {sheet}"])
        
        # 查找锚点
        col_idx = _col_to_index(column) - 1
        series = df.iloc[:, col_idx]
        mask = series.astype(str).str.strip() == text
        rows = mask[mask].index.tolist()
        
        if len(rows) < occurrence:
            return ExtractResult([], [], [f"未找到锚点文本 '{text}' 第{occurrence}次出现"])
        
        # 计算起始位置
        anchor_row = rows[occurrence - 1] + 1
        start_row = anchor_row + offset[0]
        
        # 简化：提取从锚点开始的10行数据
        end_row = min(start_row + 10, len(df))
        start_col = 1
        end_col = min(10, len(df.columns))
        
        data_df = df.iloc[start_row-1:end_row, start_col-1:end_col]
        data = data_df.values.tolist()
        
        return ExtractResult(data)
        
    except Exception as e:
        return ExtractResult([], [], [f"锚点提取失败: {str(e)}"])


# ============================================================================
# 工具函数
# ============================================================================

def _a1_to_rc(a1: str) -> tuple:
    """A1格式转换为行列坐标"""
    import re
    match = re.match(r'([A-Z]+)(\d+)', a1)
    if not match:
        raise ValueError(f"无效的A1格式: {a1}")
    
    col_str, row_str = match.groups()
    col = _col_to_index(col_str)
    row = int(row_str)
    return row, col


def _col_to_index(col_str: str) -> int:
    """列字母转换为索引（A=1, B=2, ...）"""
    result = 0
    for char in col_str:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


# ============================================================================
# 便捷函数 - 常用模式
# ============================================================================

def extract_table(file_path: str, sheet: str, start_cell: str = "A1") -> ExtractResult:
    """提取完整表格 - 自动检测边界
    
    Args:
        file_path: Excel文件路径
        sheet: 工作表名称
        start_cell: 起始单元格，默认A1
    
    Returns:
        ExtractResult: 提取结果
    """
    # 获取最后一行和最后一列
    last_row = last_data_row(file_path, sheet, 1)
    if last_row == 0:
        return ExtractResult([], [], ["表格为空"])
    
    # 简化：假设最多到Z列
    end_cell = f"Z{last_row}"
    area = f"{start_cell}:{end_cell}"
    
    return extract_simple(file_path, sheet, area)


def extract_list(file_path: str, sheet: str, column: str, start_row: int = 1) -> ExtractResult:
    """提取单列列表数据
    
    Args:
        file_path: Excel文件路径
        sheet: 工作表名称
        column: 列字母，如"A", "B"
        start_row: 起始行，默认1
    
    Returns:
        ExtractResult: 提取结果
    """
    last_row = last_data_row(file_path, sheet, _col_to_index(column))
    if last_row == 0:
        return ExtractResult([], [], ["列表为空"])
    
    area = f"{column}{start_row}:{column}{last_row}"
    return extract_simple(file_path, sheet, area)
