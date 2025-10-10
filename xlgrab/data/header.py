"""
表头处理模块

提供DataFrame表头处理功能
"""

import pandas as pd
import numpy as np
import re
from typing import Any, Optional, Union, List, Dict, Callable


def apply_header(
    df,
    header: Union[bool, int, List[int], List[str], pd.DataFrame, pd.Series] = True,
    header_join: Optional[str] = "_",
    inplace: bool = True,
):
    """
    使用本 DataFrame 顶部若干行作为列名。

    参数：
      - header: True 表示首行；整数 N 表示前 N 行；False 返回拷贝不做处理
      - header_join: 当 N>1 时，若提供分隔符则将多行头按分隔符合并为单层列；
                     否则生成 MultiIndex 多级列
      - inplace: 是否直接修改原 DataFrame，默认为 False

    返回：
      - 如果 inplace=True，返回 None（直接修改原 DataFrame）
      - 如果 inplace=False，返回处理后的新 DataFrame
    """
    # 通用：构造清洗与去重函数
    def _safe_name(val: Any) -> str:
        val = str(val)
        # 如果为空字符串或只包含空白字符，返回空字符串（后续会由_dedup_names处理）
        if not val or val.isspace():
            return ""
        val = re.sub(r'[-:/\\()\[\].,;:：；（）()【】{}·\s]', '_', val)
        val = re.sub(r'_+', '_', val)
        val = val.strip('_')
        return val

    def _dedup_names(names: List[str]) -> List[str]:
        seen: Dict[str, int] = {}
        result: List[str] = []
        for name in names:
            # 如果为空字符串，使用占位列名
            if not name:
                base = "_"
                count = seen.get(base, 0)
                if count == 0:
                    result.append(f"{base}1")  # 第一个空字符串变成 _1
                else:
                    result.append(f"{base}{count + 1}")  # 后续变成 _2, _3, ...
                seen[base] = count + 1
            else:
                base = name
                count = seen.get(base, 0)
                if count == 0:
                    result.append(base)
                else:
                    result.append(f"{base}_{count}")
                seen[base] = count + 1
        return result

    def _generate_placeholder_names(num_cols: int) -> List[str]:
        """生成占位列名 N_1, N_2, N_3, ..."""
        return [f"N_{i+1}" for i in range(num_cols)]

    # 1) header 为列表[str]/Series：直接作为列名
    if isinstance(header, (list, tuple, pd.Series)):
        # 检查是否为空或非整数类型
        is_empty = len(header) == 0
        is_non_integer = len(header) > 0 and not isinstance(next(iter(header)), (int, np.integer))
        if is_empty or is_non_integer:
            names = list(header)
            # 如果列表为空，使用占位列名
            if len(names) == 0:
                names = _generate_placeholder_names(len(df.columns))
            elif len(names) != len(df.columns):
                raise ValueError(f"提供的列名数量为 {len(names)}，与 DataFrame 列数 {len(df.columns)} 不一致")
        cleaned = [_safe_name(x) for x in names]
        cleaned = _dedup_names(cleaned)
        if inplace:
            df.columns = cleaned
            df.reset_index(drop=True, inplace=True)
            return None
        else:
            out = df.copy()
            out.columns = cleaned
            out.reset_index(drop=True, inplace=True)
            return out

    # 2) header 为 DataFrame：按多行表头合并
    if isinstance(header, pd.DataFrame):
        header.ffill(axis=1, inplace=True)
        header_block = header.astype("string")
        n = len(header_block)
        # 如果DataFrame为空，使用占位列名
        if n < 1:
            placeholder_names = _generate_placeholder_names(len(df.columns))
            if inplace:
                df.columns = placeholder_names
                df.reset_index(drop=True, inplace=True)
                return None
            else:
                out = df.copy()
                out.columns = placeholder_names
                out.reset_index(drop=True, inplace=True)
                return out
        # 从当前 df 全量返回（不丢行），仅重命名
        arrays = [header_block.iloc[i].tolist() for i in range(n)]
        if header_join is None:
            tuples = [tuple(items) for items in zip(*arrays)]
            new_columns = pd.MultiIndex.from_tuples(tuples)
        else:
            # 合并为单层列
            new_headers: List[str] = []
            for col_idx, items in enumerate(zip(*arrays)):
                values = [x for x in items if pd.notna(x)]
                unique_ordered = list(dict.fromkeys(values))
                merged = header_join.join(map(str, unique_ordered)) if unique_ordered else str(df.columns[col_idx])
                new_headers.append(_safe_name(merged))
            new_columns = _dedup_names(new_headers)
        
        if inplace:
            df.columns = new_columns
            df.reset_index(drop=True, inplace=True)
            return None
        else:
            out = df.copy()
            out.columns = new_columns
            out.reset_index(drop=True, inplace=True)
            return out

    # 3) header 为 False：不处理
    if header is False:
        if inplace:
            return None
        else:
            return df.copy()

    # 4) 与 pandas read_csv 语义对齐：
    #    - header=True 等价于 header=0（使用第0行做表头）
    #    - header=int 使用该"0基"行作为表头
    #    - header=[i,j] 使用多行表头
    if header is True:
        header = 0

    # 多行表头：list[int]
    if isinstance(header, (list, tuple)) and len(header) > 0 and isinstance(next(iter(header)), (int, np.integer)):
        idxs = list(header)
        if min(idxs) < 0 or max(idxs) >= len(df):
            raise ValueError("header 行索引超出范围")
        header_block = df.iloc[idxs, :].astype("string")
        data_start = max(idxs) + 1
        data_block = df.iloc[data_start:, :].copy()
        arrays = [header_block.iloc[i - idxs[0]].tolist() for i in idxs]
        if header_join is None:
            tuples = [tuple(items) for items in zip(*arrays)]
            data_block.columns = pd.MultiIndex.from_tuples(tuples)
        else:
            new_headers: List[str] = []
            for col_idx, items in enumerate(zip(*arrays)):
                values = [x for x in items if pd.notna(x)]
                unique_ordered = list(dict.fromkeys(values))
                merged = header_join.join(map(str, unique_ordered)) if unique_ordered else str(df.columns[col_idx])
                new_headers.append(_safe_name(merged))
            data_block.columns = _dedup_names(new_headers)
        data_block.reset_index(drop=True, inplace=True)
        if inplace:
            # 对于多行表头，需要替换整个 DataFrame
            df.__init__(data_block)
            return None
        else:
            return data_block

    # 单行表头：int
    if isinstance(header, (int, np.integer)):
        row_idx = int(header)
        if row_idx < 0 or row_idx >= len(df):
            raise ValueError("header 行索引超出范围")
        hdr = df.iloc[row_idx, :].astype("string")
        data_block = df.iloc[row_idx + 1:, :].copy()
        new_cols = [ _safe_name(x) for x in hdr.tolist() ]
        data_block.columns = _dedup_names(new_cols)
        data_block.reset_index(drop=True, inplace=True)
        if inplace:
            # 对于单行表头，需要替换整个 DataFrame
            df.__init__(data_block)
            return None
        else:
            return data_block

    # 其余类型不支持
    raise ValueError("不支持的 header 类型。支持: False/True/int/list[int]/list[str]/Series/DataFrame")
