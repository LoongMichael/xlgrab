"""
数据查找模块

提供DataFrame和Series的查找功能
"""

import pandas as pd
import numpy as np
import re
from typing import Any, Optional, Union, List, Dict, Callable


def find_idx_dataframe(
    df,
    target: Union[str, int],
    q: Union[str, re.Pattern],
    mode: str = "exact",
    na: bool = False,
    flags: int = 0,
    nth: Optional[int] = 1,
    axis: str = "column",
):
    """
    在DataFrame中查找位置：
      - 若 nth 为 None：返回所有命中位置的 ndarray
      - 若 nth 为正整数：返回第 n 次命中的位置（int），未命中返回 -1

    mode:
      - 'exact'    : 等值匹配（默认 & 最快）。基于底层 ndarray 等值比较，加速且省内存。
      - 'contains' : 字面子串匹配（regex=False）。避免正则元字符带来的歧义。
      - 'regex'    : 正则匹配（可用 flags 传 re.IGNORECASE 等）。当 q 为 Pattern 时通常以其自带 flags 为准。

    参数：
      - target: 要搜索的目标（列名str或行索引int）。
      - q: 查询（str 或 re.Pattern）。
      - mode: 匹配模式，见上。
      - na: 仅用于 contains/regex，缺失值的布尔值（传入 .str.contains 的 na）。
      - flags: 仅用于 regex，正则标志位。
      - nth: 选择第几次命中。None 返回全部；>0 从头数；<0 从尾数（-1 为最后一次）。
      - axis: 搜索轴，"column"（按列搜索）或 "row"（按行搜索）。

    说明：
      - 若未命中且 nth 非 None，返回 -1。
      - 输入无效模式会抛 ValueError。
    """
    # 自动判断axis：如果 target 是整数且 axis 为 column，则允许以列索引方式查找
    if isinstance(target, int) and axis == "column":
        if 0 <= target < len(df.columns):
            column_name = df.columns[target]
            column_data = df[column_name]
            return find_idx_series(column_data, q, mode=mode, na=na, flags=flags, nth=nth)
        else:
            # 列索引越界时回退为按行搜索
            axis = "row"
    
    if axis == "column":
        # 按列搜索
        if target not in df.columns:
            raise ValueError(f"列 '{target}' 不存在")
        column_data = df[target]
        return find_idx_series(column_data, q, mode=mode, na=na, flags=flags, nth=nth)
    
    elif axis == "row":
        # 按行搜索 - 复用Series的find_idx方法
        if target not in df.index:
            raise ValueError(f"行索引 '{target}' 不存在")
        
        # 获取指定行的数据
        row_data = df.loc[target]
        return find_idx_series(row_data, q, mode=mode, na=na, flags=flags, nth=nth)
    
    else:
        raise ValueError("axis must be 'column' or 'row'")


def find_idx_series(
    series,
    q: Union[str, re.Pattern],
    mode: str = "exact",
    na: bool = False,
    flags: int = 0,
    nth: Optional[int] = 1,
):
    """
    返回命中位置：
      - 若 nth 为 None：返回所有命中位置的 ndarray
      - 若 nth 为正整数：返回第 n 次命中的位置（int），未命中返回 -1

    mode:
      - 'exact'    : 等值匹配（默认 & 最快）。基于底层 ndarray 等值比较，加速且省内存。
      - 'contains' : 字面子串匹配（regex=False）。避免正则元字符带来的歧义。
      - 'regex'    : 正则匹配（可用 flags 传 re.IGNORECASE 等）。当 q 为 Pattern 时通常以其自带 flags 为准。

    参数：
      - q: 查询（str 或 re.Pattern）。
      - mode: 匹配模式，见上。
      - na: 仅用于 contains/regex，缺失值的布尔值（传入 .str.contains 的 na）。
      - flags: 仅用于 regex，正则标志位。
      - nth: 选择第几次命中。None 返回全部；>0 从头数；<0 从尾数（-1 为最后一次）。

    说明：
      - 若未命中且 nth 非 None，返回 -1。
      - 输入无效模式会抛 ValueError。
    """
    
    # exact：使用底层 ndarray 做等值比较，性能最优
    if mode == "exact":
        arr = series.to_numpy(copy=False)
        if pd.isna(q):
            # 查找缺失值位置（覆盖 None/np.nan/pd.NA 等）
            idx = np.flatnonzero(pd.isna(arr))
        else:
            idx = np.flatnonzero(arr == q)
    elif mode == "contains":
        # contains：字面子串匹配（regex=False），避免正则引擎开销与语义歧义
        arr = series.astype("string")
        mask = arr.str.contains(str(q), regex=False, na=na)
        idx = np.flatnonzero(mask.to_numpy())
    elif mode == "regex":
        # regex：正则匹配，可通过 flags 控制大小写等
        arr = series.astype("string")
        mask = arr.str.contains(q, regex=True, na=na, flags=flags)
        idx = np.flatnonzero(mask.to_numpy())
    else:
        raise ValueError("mode must be 'exact' | 'contains' | 'regex'")

    # 命中次序选择：None → 全部；>0 → 第 n 个；<0 → 从尾部计数
    if nth is None:
        return idx
    if nth == 0:
        raise ValueError("nth must be a non-zero integer or None")
    
    # 确保nth是整数类型
    if not isinstance(nth, int):
        raise ValueError("nth must be an integer or None")

    k = nth - 1 if nth > 0 else idx.size + nth
    return int(idx[k]) if 0 <= k < idx.size else -1
