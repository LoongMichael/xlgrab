"""
Excel 范围操作模块

提供Excel范围读取、偏移选择、DSL选择等功能
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


def excel_range(
    df,
    *ranges,
    header: bool = True,
    index_col: Optional[Union[int, str]] = None,
):
    """
    将Excel数据区间转换为DataFrame的数据区间，支持多个区域合并
    
    参数：
      - *ranges: Excel数据区间，支持单个单元格（如 'B2'）或区域（如 'B2:D6'），
                也支持多个区域，如 'B2:D6,K9:L11,K13:L15'
      - header: 是否将第一行作为列名
      - index_col: 指定作为索引的列（列名或列索引）
    
    返回：
      - DataFrame: 转换后的DataFrame
    
    注意：
      - 如果请求的区域超出DataFrame的实际范围，会自动裁剪到有效边界并发出警告
      - 起始位置必须在有效范围内，否则会抛出错误
    
    示例：
      df.excel_range('B2')  # 获取B2单元格的数据
      df.excel_range('B2:D6')  # 获取B2到D6的数据
      df.excel_range('A1:C5', header=True)  # 第一行作为列名
      df.excel_range('A1:C5', header=True, index_col=0)  # 第一列作为索引
      df.excel_range('B2:D6', 'K9:L11', 'K13:L15')  # 合并多个区域
      df.excel_range('B2:Z1000')  # 如果超出范围，自动裁剪到实际最大行列
    """
    if not ranges:
        raise ValueError("至少需要提供一个Excel区间")
    
    # 处理多个区域
    all_dfs = []
    
    for range_str in ranges:
        # 解析Excel区间
        if ',' in range_str:
            # 处理多个区间，如 'B2:D6,K9:L11'
            sub_ranges = range_str.split(',')
            for sub_range in sub_ranges:
                range_df = _parse_excel_range(df, sub_range.strip())
                all_dfs.append(range_df)
        else:
            # 单个区间
            range_df = _parse_excel_range(df, range_str)
            all_dfs.append(range_df)
    
    # 合并所有区域
    if len(all_dfs) == 1:
        result_df = all_dfs[0]
    else:
        # 垂直合并多个区域
        result_df = pd.concat(all_dfs, ignore_index=True)
    
    # 处理header
    if header and len(result_df) > 0:
        # 将第一行作为列名
        new_columns = result_df.iloc[0].tolist()
        result_df = result_df.iloc[1:].copy()
        result_df.columns = new_columns
        result_df.reset_index(drop=True, inplace=True)
    
    # 处理index_col
    if index_col is not None and len(result_df) > 0:
        if isinstance(index_col, str):
            if index_col in result_df.columns:
                result_df.set_index(index_col, inplace=True)
            else:
                raise ValueError(f"列名 '{index_col}' 不存在")
        elif isinstance(index_col, int):
            if 0 <= index_col < len(result_df.columns):
                result_df.set_index(result_df.columns[index_col], inplace=True)
            else:
                raise ValueError(f"列索引 {index_col} 超出范围")

    return result_df


def _parse_excel_range(df, range_str: str):
    """解析Excel区间字符串，如 'B2:D6' 或单个单元格 'B2'
    
    如果请求的区域超出 DataFrame 的实际范围，会自动裁剪到有效边界。
    """
    if not OPENPYXL_AVAILABLE:
        raise ImportError("需要安装 openpyxl 库来解析Excel区间")
    
    try:
        # 检查是否是单个单元格（不包含冒号）
        if ':' not in range_str:
            # 单个单元格，转换为范围格式 "B2:B2"
            range_str = f"{range_str}:{range_str}"
        
        # 使用openpyxl解析区间
        from openpyxl.utils import range_boundaries
        min_col, min_row, max_col, max_row = range_boundaries(range_str)
        
        # 转换为pandas索引（从0开始）
        start_row_idx = min_row - 1
        end_row_idx = max_row - 1
        start_col_idx = min_col - 1
        end_col_idx = max_col - 1
        
        # 获取 DataFrame 的实际边界
        df_max_row = len(df) - 1
        df_max_col = len(df.columns) - 1
        
        # 检查起始位置是否完全超出范围
        if start_row_idx > df_max_row or start_col_idx > df_max_col:
            raise ValueError(f"起始位置超出范围: 请求行{min_row}列{min_col}，但DataFrame只有{len(df)}行{len(df.columns)}列")
        
        if start_row_idx < 0 or start_col_idx < 0:
            raise ValueError(f"起始位置无效: 行{min_row}列{min_col}必须大于0")
        
        # 自动裁剪结束位置到有效范围
        original_end_row = end_row_idx
        original_end_col = end_col_idx
        
        end_row_idx = min(end_row_idx, df_max_row)
        end_col_idx = min(end_col_idx, df_max_col)
        
        # 如果发生了裁剪，发出警告
        if end_row_idx < original_end_row or end_col_idx < original_end_col:
            warnings.warn(
                f"请求的区域超出DataFrame范围，已自动裁剪: "
                f"请求到第{max_row}行第{max_col}列，实际返回到第{end_row_idx+1}行第{end_col_idx+1}列",
                UserWarning
            )
        
        # 获取数据区间
        # Excel区间是包含边界的，所以需要+1
        range_df = df.iloc[start_row_idx:end_row_idx+1, start_col_idx:end_col_idx+1].copy()
        return range_df
        
    except Exception as e:
        raise ValueError(f"无法解析Excel区间 '{range_str}': {e}")


def offset_range(
    df,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
    offset_rows: int = 0,
    offset_cols: int = 0,
    offset_start_row: Optional[int] = None,
    offset_end_row: Optional[int] = None,
    offset_start_col: Optional[int] = None,
    offset_end_col: Optional[int] = None,
    clip_to_bounds: bool = False,
):
    """
    基于Excel行列坐标和偏移量获取数据区间，支持统一偏移和分别偏移两种模式
    
    参数：
      - start_row: 起始行（从1开始）
      - end_row: 结束行（从1开始）
      - start_col: 起始列（从1开始，A=1, B=2, C=3...）
      - end_col: 结束列（从1开始，A=1, B=2, C=3...）
      - offset_rows: 行偏移量（正数向下，负数向上）- 统一偏移模式
      - offset_cols: 列偏移量（正数向右，负数向左）- 统一偏移模式
      - offset_start_row: 起始行偏移量（分别偏移模式）
      - offset_end_row: 结束行偏移量（分别偏移模式）
      - offset_start_col: 起始列偏移量（分别偏移模式）
      - offset_end_col: 结束列偏移量（分别偏移模式）
      - clip_to_bounds: 是否自动裁剪到有效范围
    
    返回：
      - DataFrame: 偏移后的数据区间
    
    示例：
      # 统一偏移模式（默认）
      df.offset_range(1, 5, 2, 6, offset_rows=2, offset_cols=-1)
      
      # 分别偏移模式
      df.offset_range(1, 5, 2, 6, offset_start_row=1, offset_end_row=2, offset_start_col=-1, offset_end_col=1)
      
      # 自动裁剪模式
      df.offset_range(1, 5, 2, 6, offset_rows=10, clip_to_bounds=True)
    """
    # 判断使用哪种偏移模式
    use_flexible = any(x is not None for x in [offset_start_row, offset_end_row, offset_start_col, offset_end_col])
    
    if use_flexible:
        # 分别偏移模式
        new_start_row = start_row + (offset_start_row or 0)
        new_end_row = end_row + (offset_end_row or 0)
        new_start_col = start_col + (offset_start_col or 0)
        new_end_col = end_col + (offset_end_col or 0)
    else:
        # 统一偏移模式
        new_start_row = start_row + offset_rows
        new_end_row = end_row + offset_rows
        new_start_col = start_col + offset_cols
        new_end_col = end_col + offset_cols
    
    # 边界处理
    if clip_to_bounds:
        # 自动裁剪到有效范围
        max_row = len(df)
        max_col = len(df.columns)
        new_start_row = max(1, min(new_start_row, max_row))
        new_end_row = max(1, min(new_end_row, max_row))
        new_start_col = max(1, min(new_start_col, max_col))
        new_end_col = max(1, min(new_end_col, max_col))
        
        # 检查区间是否有效
        if new_start_row > new_end_row or new_start_col > new_end_col:
            raise ValueError("偏移后的区间无效")
    else:
        # 严格检查，超出范围抛异常
        if new_start_row < 1 or new_end_row > len(df):
            raise ValueError(f"偏移后行索引超出范围: {new_start_row}-{new_end_row}，DataFrame只有{len(df)}行")
        
        if new_start_col < 1 or new_end_col > len(df.columns):
            raise ValueError(f"偏移后列索引超出范围: {new_start_col}-{new_end_col}，DataFrame只有{len(df.columns)}列")
    
    # 将Excel坐标转换为pandas索引（从0开始）
    start_row_idx = new_start_row - 1
    end_row_idx = new_end_row - 1
    start_col_idx = new_start_col - 1
    end_col_idx = new_end_col - 1
    
    # 获取偏移后的数据区间
    result_df = df.iloc[start_row_idx:end_row_idx+1, start_col_idx:end_col_idx+1].copy()
    
    return result_df


def select_range(
    df,
    *,
    start: Optional[Union[str, int, tuple, list]] = None,
    end: Optional[Union[str, int, tuple, list]] = None,
    start_row: Optional[Union[str, int, tuple, list]] = None,
    end_row: Optional[Union[str, int, tuple, list]] = None,
    start_col: Optional[Union[str, int, tuple, list]] = None,
    end_col: Optional[Union[str, int, tuple, list]] = None,
    clip: bool = True,
    # 偏移相关参数（与 offset_range 一致的两种模式：统一偏移/分别偏移）
    offset_rows: int = 0,
    offset_cols: int = 0,
    offset_start_row: Optional[int] = None,
    offset_end_row: Optional[int] = None,
    offset_start_col: Optional[int] = None,
    offset_end_col: Optional[int] = None,
):
    """
    DSL风格的区间选择，优雅表达混合场景，最终构建 iloc 切片。

    支持的规格（spec）写法：
      - 字符串：
        * "A2"         → 单元格（同时指定行与列）
        * "F"/"AA"     → 列（Excel 列字母）
      - 整数：按 Excel 习惯 1 基行/列索引（内部转 0 基）
      - 元组/列表：
        * ("cell", "A2")
        * ("row", 10)
        * ("col", "F" | 6)
        * ("find-row", target, q, {mode, nth, na, flags})
        * ("find-col", target, q, {mode, nth, na, flags})

    参数优先级：
      - start/end 可一次性给端点；
      - start_row/start_col/end_row/end_col 可覆盖对应维度；
      - 未指定的边界使用默认：start_row=1, start_col=1, end_row=末行, end_col=末列。
        因此通常无需显式写 'end'：例如仅给出 `start='B2'` 即表示从 B2 一直到表尾；
        仅给 `start_row` 或 `start_col` 也分别表示到末行或末列。

    例：
      df.select_range(start='B2')                       # 从 B2 到末行末列
      df.select_range(start_row=('row', 3))             # 第 3 行到末行，列为全列
      df.select_range(start_col=('col', 'C'))           # 第 C 列到末列，行为全行
    """

    num_rows = len(df)
    num_cols = len(df.columns)

    def is_cell_str(s: str) -> bool:
        return isinstance(s, str) and any(c.isdigit() for c in s) and any(c.isalpha() for c in s)

    def excel_col_to_idx(col_label: str) -> int:
        # A→1, B→2 ... 转 1 基
        col_label = col_label.strip().upper()
        val = 0
        for ch in col_label:
            if not ('A' <= ch <= 'Z'):
                raise ValueError(f"无效的列标记: {col_label}")
            val = val * 26 + (ord(ch) - ord('A') + 1)
        return val

    def normalize_1based_idx(v: int, upper: int, default_end: bool = False) -> int:
        # 输入 1 基整数，返回 0 基索引（限定边界；default_end 为 True 表示默认末端）
        if v is None:
            return (upper - 1) if default_end else 0
        v1 = int(v)
        if clip:
            v1 = max(1, min(v1, upper))
        else:
            if v1 < 1 or v1 > upper:
                raise ValueError(f"索引 {v1} 超出范围 [1, {upper}]")
        return v1 - 1  # 转 0 基

    def parse_row_spec(spec, default_end: bool = False) -> Optional[int]:
        if spec is None:
            return None
        if isinstance(spec, int):
            return normalize_1based_idx(spec, num_rows, default_end)
        if isinstance(spec, str):
            if is_cell_str(spec):
                # 解析单元格字符串，如 "A2"
                if not OPENPYXL_AVAILABLE:
                    raise ImportError("需要安装 openpyxl 库来解析单元格字符串")
                try:
                    col, row = coordinate_to_tuple(spec)
                    return normalize_1based_idx(row, num_rows, default_end)
                except Exception as e:
                    raise ValueError(f"无法解析单元格字符串 '{spec}': {e}")
            else:
                # 纯列字母，不支持
                raise ValueError(f"字符串 '{spec}' 不是有效的单元格格式")
        if isinstance(spec, (tuple, list)) and len(spec) >= 2:
            spec_type = spec[0]
            if spec_type == "row":
                return normalize_1based_idx(spec[1], num_rows, default_end)
            elif spec_type == "cell":
                if not OPENPYXL_AVAILABLE:
                    raise ImportError("需要安装 openpyxl 库来解析单元格字符串")
                try:
                    col, row = coordinate_to_tuple(spec[1])
                    return normalize_1based_idx(row, num_rows, default_end)
                except Exception as e:
                    raise ValueError(f"无法解析单元格字符串 '{spec[1]}': {e}")
            elif spec_type == "find-row":
                # 需要调用 find_idx 方法
                from ..data.search import find_idx_dataframe
                target, q = spec[1], spec[2]
                opts = spec[3] if len(spec) > 3 else {}
                mode = opts.get("mode", "exact")
                na = opts.get("na", False)
                flags = opts.get("flags", 0)
                nth = opts.get("nth", 1)
                pos = find_idx_dataframe(df, target, q, mode=mode, na=na, flags=flags, nth=nth, axis="column")
                if isinstance(pos, np.ndarray):
                    pos = int(pos[0]) if pos.size > 0 else -1
                return None if pos is None or pos < 0 else int(pos)
        raise ValueError(f"不支持的行规格: {spec}")

    def parse_col_spec(spec, default_end: bool = False) -> Optional[int]:
        if spec is None:
            return None
        if isinstance(spec, int):
            return normalize_1based_idx(spec, num_cols, default_end)
        if isinstance(spec, str):
            if is_cell_str(spec):
                # 解析单元格字符串，如 "A2"
                if not OPENPYXL_AVAILABLE:
                    raise ImportError("需要安装 openpyxl 库来解析单元格字符串")
                try:
                    col, row = coordinate_to_tuple(spec)
                    return normalize_1based_idx(col, num_cols, default_end)
                except Exception as e:
                    raise ValueError(f"无法解析单元格字符串 '{spec}': {e}")
            else:
                # 列字母，如 "F", "AA"
                col_idx = excel_col_to_idx(spec)
                return normalize_1based_idx(col_idx, num_cols, default_end)
        if isinstance(spec, (tuple, list)) and len(spec) >= 2:
            spec_type = spec[0]
            if spec_type == "col":
                col_val = spec[1]
                if isinstance(col_val, str):
                    col_idx = excel_col_to_idx(col_val)
                else:
                    col_idx = col_val
                return normalize_1based_idx(col_idx, num_cols, default_end)
            elif spec_type == "cell":
                if not OPENPYXL_AVAILABLE:
                    raise ImportError("需要安装 openpyxl 库来解析单元格字符串")
                try:
                    col, row = coordinate_to_tuple(spec[1])
                    return normalize_1based_idx(col, num_cols, default_end)
                except Exception as e:
                    raise ValueError(f"无法解析单元格字符串 '{spec[1]}': {e}")
            elif spec_type == "find-col":
                # 需要调用 find_idx 方法
                from ..data.search import find_idx_dataframe
                target, q = spec[1], spec[2]
                opts = spec[3] if len(spec) > 3 else {}
                mode = opts.get("mode", "exact")
                na = opts.get("na", False)
                flags = opts.get("flags", 0)
                nth = opts.get("nth", 1)
                pos = find_idx_dataframe(df, target, q, mode=mode, na=na, flags=flags, nth=nth, axis="row")
                if isinstance(pos, np.ndarray):
                    pos = int(pos[0]) if pos.size > 0 else -1
                return None if pos is None or pos < 0 else int(pos)
        raise ValueError(f"不支持的列规格: {spec}")

    # 解析 start/end
    start_r_from_start = start_c_from_start = None
    end_r_from_end = end_c_from_end = None

    if start is not None:
        if isinstance(start, str) and is_cell_str(start):
            if not OPENPYXL_AVAILABLE:
                raise ImportError("需要安装 openpyxl 库来解析单元格字符串")
            try:
                col, row = coordinate_to_tuple(start)
                start_r_from_start = normalize_1based_idx(row, num_rows)
                start_c_from_start = normalize_1based_idx(col, num_cols)
            except Exception as e:
                raise ValueError(f"无法解析起始单元格 '{start}': {e}")
        else:
            raise ValueError(f"start 参数必须是单元格字符串，如 'A2'，当前为: {start}")

    if end is not None:
        if isinstance(end, str) and is_cell_str(end):
            if not OPENPYXL_AVAILABLE:
                raise ImportError("需要安装 openpyxl 库来解析单元格字符串")
            try:
                col, row = coordinate_to_tuple(end)
                end_r_from_end = normalize_1based_idx(row, num_rows, default_end=True)
                end_c_from_end = normalize_1based_idx(col, num_cols, default_end=True)
            except Exception as e:
                raise ValueError(f"无法解析结束单元格 '{end}': {e}")
        else:
            raise ValueError(f"end 参数必须是单元格字符串，如 'Z100'，当前为: {end}")

    # 合并优先级：显式的 start_row/start_col/end_row/end_col 覆盖 start/end 推断
    row_start_idx = parse_row_spec(start_row) if start_row is not None else start_r_from_start
    col_start_idx = parse_col_spec(start_col) if start_col is not None else start_c_from_start
    row_end_idx = parse_row_spec(end_row, default_end=True) if end_row is not None else end_r_from_end
    col_end_idx = parse_col_spec(end_col, default_end=True) if end_col is not None else end_c_from_end

    # 应用偏移
    if any(x is not None for x in [offset_start_row, offset_end_row, offset_start_col, offset_end_col]):
        # 分别偏移模式
        if row_start_idx is not None:
            row_start_idx += (offset_start_row or 0)
        if row_end_idx is not None:
            row_end_idx += (offset_end_row or 0)
        if col_start_idx is not None:
            col_start_idx += (offset_start_col or 0)
        if col_end_idx is not None:
            col_end_idx += (offset_end_col or 0)
    else:
        # 统一偏移模式
        if row_start_idx is not None:
            row_start_idx += offset_rows
        if row_end_idx is not None:
            row_end_idx += offset_rows
        if col_start_idx is not None:
            col_start_idx += offset_cols
        if col_end_idx is not None:
            col_end_idx += offset_cols

    # 边界处理
    if clip:
        row_start_idx = max(0, min(row_start_idx or 0, num_rows - 1))
        row_end_idx = max(0, min(row_end_idx or (num_rows - 1), num_rows - 1))
        col_start_idx = max(0, min(col_start_idx or 0, num_cols - 1))
        col_end_idx = max(0, min(col_end_idx or (num_cols - 1), num_cols - 1))
    else:
        if row_start_idx is not None and (row_start_idx < 0 or row_start_idx >= num_rows):
            raise ValueError(f"起始行索引 {row_start_idx} 超出范围 [0, {num_rows-1}]")
        if row_end_idx is not None and (row_end_idx < 0 or row_end_idx >= num_rows):
            raise ValueError(f"结束行索引 {row_end_idx} 超出范围 [0, {num_rows-1}]")
        if col_start_idx is not None and (col_start_idx < 0 or col_start_idx >= num_cols):
            raise ValueError(f"起始列索引 {col_start_idx} 超出范围 [0, {num_cols-1}]")
        if col_end_idx is not None and (col_end_idx < 0 or col_end_idx >= num_cols):
            raise ValueError(f"结束列索引 {col_end_idx} 超出范围 [0, {num_cols-1}]")

    # 确保 start <= end
    if row_start_idx is not None and row_end_idx is not None and row_start_idx > row_end_idx:
        row_start_idx, row_end_idx = row_end_idx, row_start_idx
    if col_start_idx is not None and col_end_idx is not None and col_start_idx > col_end_idx:
        col_start_idx, col_end_idx = col_end_idx, col_start_idx

    # 构建切片
    row_slice = slice(row_start_idx, (row_end_idx + 1) if row_end_idx is not None else None)
    col_slice = slice(col_start_idx, (col_end_idx + 1) if col_end_idx is not None else None)

    return df.iloc[row_slice, col_slice].copy()
