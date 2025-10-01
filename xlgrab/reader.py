from typing import Optional, Any
import pandas as pd

"""
极简的 Excel 读取器
- 参考pandas设计，提供一步调用方式
- 专注核心功能，简洁高效
"""


def get_sheet(file_path: str, sheet_name: str) -> Optional[pd.DataFrame]:
    """获取sheet的DataFrame，不存在返回None"""
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine="calamine")
    except Exception:
        return None


def get_region(file_path: str, sheet_name: str, start_row: int, end_row: int, start_col: int, end_col: int) -> Optional[pd.DataFrame]:
    """获取指定区域的数据"""
    df = get_sheet(file_path, sheet_name)
    if df is None:
        return None
    
    return df.iloc[start_row-1:end_row, start_col-1:end_col]


def get_cell(file_path: str, sheet_name: str, row: int, col: int) -> Any:
    """获取单个单元格值"""
    region = get_region(file_path, sheet_name, row, row, col, col)
    if region is None or region.empty:
        return None
    value = region.iloc[0, 0]
    return None if pd.isna(value) else value


def last_data_row(file_path: str, sheet_name: str, col: int = 1) -> int:
    """获取指定列的最后数据行"""
    df = get_sheet(file_path, sheet_name)
    if df is None:
        return 0
    
    series = df.iloc[:, col-1]
    mask = ~series.isna() & (series.astype(str).str.strip() != "")
    idx = mask[mask].index
    
    return int(idx[-1]) + 1 if len(idx) > 0 else 0


