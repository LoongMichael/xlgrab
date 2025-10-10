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
    
    # ==================== 数据探索方法 ====================
    
    def excel_range(
        self,
        *ranges,
        header: bool = True,
        index_col: Optional[Union[int, str]] = None,
    ):
        """
        将Excel数据区间转换为DataFrame的数据区间，支持多个区域合并
        
        参数：
          - *ranges: Excel数据区间，支持多个区域，如 'B2:D6' 或 'B2:D6,K9:L11,K13:L15'
          - header: 是否将第一行作为列名
          - index_col: 指定作为索引的列（列名或列索引）
        
        返回：
          - DataFrame: 转换后的DataFrame
        
        示例：
          df.excel_range('B2:D6')  # 获取B2到D6的数据
          df.excel_range('A1:C5', header=True)  # 第一行作为列名
          df.excel_range('A1:C5', header=True, index_col=0)  # 第一列作为索引
          df.excel_range('B2:D6', 'K9:L11', 'K13:L15')  # 合并多个区域
        """
        if not ranges:
            raise ValueError("至少需要提供一个Excel区间")
        
        # 处理多个区域
        all_dfs = []
        
        for range_str in ranges:
            # 解析Excel区间
            if ':' not in range_str:
                raise ValueError(f"无效的Excel区间格式: {range_str}。请使用格式如 'B2:D6'")
            
            start_cell, end_cell = range_str.split(':')
            
            # 使用openpyxl解析坐标
            start_row, start_col = coordinate_to_tuple(start_cell.upper())
            end_row, end_col = coordinate_to_tuple(end_cell.upper())
            
            # 转换为pandas索引（从0开始）
            start_row_idx = start_row - 1
            end_row_idx = end_row - 1
            start_col_idx = start_col - 1
            end_col_idx = end_col - 1
            
            # 检查索引是否在DataFrame范围内
            if start_row_idx < 0 or end_row_idx >= len(self):
                raise ValueError(f"行索引超出范围: {start_row}-{end_row}，DataFrame只有{len(self)}行")
            
            if start_col_idx < 0 or end_col_idx >= len(self.columns):
                raise ValueError(f"列索引超出范围: {start_col}-{end_col}，DataFrame只有{len(self.columns)}列")
            
            # 获取数据区间
            # Excel区间是包含边界的，所以需要+1
            range_df = self.iloc[start_row_idx:end_row_idx+1, start_col_idx:end_col_idx+1].copy()
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
    
    def offset_range(
        self,
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
            max_row = len(self)
            max_col = len(self.columns)
            new_start_row = max(1, min(new_start_row, max_row))
            new_end_row = max(1, min(new_end_row, max_row))
            new_start_col = max(1, min(new_start_col, max_col))
            new_end_col = max(1, min(new_end_col, max_col))
            
            # 检查区间是否有效
            if new_start_row > new_end_row or new_start_col > new_end_col:
                raise ValueError("偏移后的区间无效")
        else:
            # 严格检查，超出范围抛异常
            if new_start_row < 1 or new_end_row > len(self):
                raise ValueError(f"偏移后行索引超出范围: {new_start_row}-{new_end_row}，DataFrame只有{len(self)}行")
            
            if new_start_col < 1 or new_end_col > len(self.columns):
                raise ValueError(f"偏移后列索引超出范围: {new_start_col}-{new_end_col}，DataFrame只有{len(self.columns)}列")
        
        # 将Excel坐标转换为pandas索引（从0开始）
        start_row_idx = new_start_row - 1
        end_row_idx = new_end_row - 1
        start_col_idx = new_start_col - 1
        end_col_idx = new_end_col - 1
        
        # 获取偏移后的数据区间
        result_df = self.iloc[start_row_idx:end_row_idx+1, start_col_idx:end_col_idx+1].copy()
        
        return result_df
    
    
    def select_range(
        self,
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

        num_rows = len(self)
        num_cols = len(self.columns)

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
                    raise ValueError(f"索引超出范围: {v1}，有效范围 1..{upper}")
            return v1 - 1

        def parse_row_spec(spec, default_end: bool = False) -> Optional[int]:
            # 返回 0 基行索引；None 表示未指定
            if spec is None:
                return None
            # 字符串
            if isinstance(spec, str):
                if spec.lower() == "end":
                    return num_rows - 1
                if is_cell_str(spec):
                    r, _c = coordinate_to_tuple(spec.upper())
                    return normalize_1based_idx(r, num_rows)
                # 纯列字母字符串不在行语境中生效
                raise ValueError(f"无法将字符串 '{spec}' 解析为行索引")
            # 整数（1 基）
            if isinstance(spec, int):
                return normalize_1based_idx(spec, num_rows, default_end=default_end)
            # 元组/列表
            if isinstance(spec, (tuple, list)) and spec:
                kind = str(spec[0]).lower()
                if kind == "cell":
                    r, _c = coordinate_to_tuple(str(spec[1]).upper())
                    return normalize_1based_idx(r, num_rows)
                if kind == "row":
                    v = spec[1]
                    if isinstance(v, str) and v.lower() == "end":
                        return num_rows - 1
                    return normalize_1based_idx(int(v), num_rows, default_end=default_end)
                if kind == "find-row":
                    # ("find-row", target, q, {mode, nth, na, flags})
                    target, q = spec[1], spec[2]
                    opts: Dict[str, Any] = spec[3] if len(spec) > 3 and isinstance(spec[3], dict) else {}
                    mode = opts.get("mode", "exact")
                    na = opts.get("na", False)
                    flags = opts.get("flags", 0)
                    nth = opts.get("nth", 1)
                    pos = XlDataFrame.find_idx(self, target, q, mode=mode, na=na, flags=flags, nth=nth, axis="column")
                    if isinstance(pos, np.ndarray):
                        pos = int(pos[0]) if pos.size > 0 else -1
                    return None if pos is None or pos < 0 else int(pos)
            raise ValueError(f"无法解析行规格: {spec}")

        def parse_col_spec(spec, default_end: bool = False) -> Optional[int]:
            # 返回 0 基列索引；None 表示未指定
            if spec is None:
                return None
            if isinstance(spec, str):
                if spec.lower() == "end":
                    return num_cols - 1
                if is_cell_str(spec):
                    _r, c = coordinate_to_tuple(spec.upper())
                    return normalize_1based_idx(c, num_cols)
                # 纯列字母
                if spec.isalpha():
                    c1 = excel_col_to_idx(spec)
                    return normalize_1based_idx(c1, num_cols, default_end=default_end)
                raise ValueError(f"无法将字符串 '{spec}' 解析为列索引")
            if isinstance(spec, int):
                return normalize_1based_idx(spec, num_cols, default_end=default_end)
            if isinstance(spec, (tuple, list)) and spec:
                kind = str(spec[0]).lower()
                if kind == "cell":
                    _r, c = coordinate_to_tuple(str(spec[1]).upper())
                    return normalize_1based_idx(c, num_cols)
                if kind == "col":
                    v = spec[1]
                    if isinstance(v, str):
                        c1 = excel_col_to_idx(v)
                        return normalize_1based_idx(c1, num_cols, default_end=default_end)
                    return normalize_1based_idx(int(v), num_cols, default_end=default_end)
                if kind == "find-col":
                    # ("find-col", targetRowIndexOrLabel, q, {mode, nth, na, flags})
                    target, q = spec[1], spec[2]
                    opts: Dict[str, Any] = spec[3] if len(spec) > 3 and isinstance(spec[3], dict) else {}
                    mode = opts.get("mode", "exact")
                    na = opts.get("na", False)
                    flags = opts.get("flags", 0)
                    nth = opts.get("nth", 1)
                    pos = XlDataFrame.find_idx(self, target, q, mode=mode, na=na, flags=flags, nth=nth, axis="row")
                    if isinstance(pos, np.ndarray):
                        pos = int(pos[0]) if pos.size > 0 else -1
                    return None if pos is None or pos < 0 else int(pos)
            raise ValueError(f"无法解析列规格: {spec}")

        # 先用 start / end 解析端点（允许是单元格，或行/列标记）
        start_r_from_start = start_c_from_start = None
        end_r_from_end = end_c_from_end = None

        if start is not None:
            if isinstance(start, str) and is_cell_str(start):
                r, c = coordinate_to_tuple(start.upper())
                start_r_from_start = normalize_1based_idx(r, num_rows)
                start_c_from_start = normalize_1based_idx(c, num_cols)
            elif isinstance(start, (tuple, list)) and start and str(start[0]).lower() == "cell":
                r, c = coordinate_to_tuple(str(start[1]).upper())
                start_r_from_start = normalize_1based_idx(r, num_rows)
                start_c_from_start = normalize_1based_idx(c, num_cols)
            else:
                # 非 cell：分别试图作为行与列去解析（允许只命中一维）
                try:
                    start_r_from_start = parse_row_spec(start)
                except Exception:
                    pass
                try:
                    start_c_from_start = parse_col_spec(start)
                except Exception:
                    pass

        if end is not None:
            if isinstance(end, str) and is_cell_str(end):
                r, c = coordinate_to_tuple(end.upper())
                end_r_from_end = normalize_1based_idx(r, num_rows)
                end_c_from_end = normalize_1based_idx(c, num_cols)
            elif isinstance(end, (tuple, list)) and end and str(end[0]).lower() == "cell":
                r, c = coordinate_to_tuple(str(end[1]).upper())
                end_r_from_end = normalize_1based_idx(r, num_rows)
                end_c_from_end = normalize_1based_idx(c, num_cols)
            else:
                try:
                    end_r_from_end = parse_row_spec(end, default_end=True)
                except Exception:
                    pass
                try:
                    end_c_from_end = parse_col_spec(end, default_end=True)
                except Exception:
                    pass

        # 合并优先级：显式的 start_row/start_col/end_row/end_col 覆盖 start/end 推断
        row_start_idx = parse_row_spec(start_row) if start_row is not None else start_r_from_start
        col_start_idx = parse_col_spec(start_col) if start_col is not None else start_c_from_start
        row_end_idx = parse_row_spec(end_row, default_end=True) if end_row is not None else end_r_from_end
        col_end_idx = parse_col_spec(end_col, default_end=True) if end_col is not None else end_c_from_end

        # 默认值兜底
        if row_start_idx is None:
            row_start_idx = 0
        if col_start_idx is None:
            col_start_idx = 0
        if row_end_idx is None:
            row_end_idx = num_rows - 1
        if col_end_idx is None:
            col_end_idx = num_cols - 1

        # 边界与顺序
        if clip:
            row_start_idx = max(0, min(row_start_idx, num_rows - 1))
            row_end_idx = max(0, min(row_end_idx, num_rows - 1))
            col_start_idx = max(0, min(col_start_idx, num_cols - 1))
            col_end_idx = max(0, min(col_end_idx, num_cols - 1))
        else:
            if not (0 <= row_start_idx < num_rows and 0 <= row_end_idx < num_rows):
                raise ValueError("行索引超出范围")
            if not (0 <= col_start_idx < num_cols and 0 <= col_end_idx < num_cols):
                raise ValueError("列索引超出范围")

        if row_start_idx > row_end_idx:
            row_start_idx, row_end_idx = row_end_idx, row_start_idx
        if col_start_idx > col_end_idx:
            col_start_idx, col_end_idx = col_end_idx, col_start_idx

        # 复用 offset_range 执行偏移与切片
        return XlDataFrame.offset_range(self,
            start_row=row_start_idx + 1,
            end_row=row_end_idx + 1,
            start_col=col_start_idx + 1,
            end_col=col_end_idx + 1,
            offset_rows=offset_rows,
            offset_cols=offset_cols,
            offset_start_row=offset_start_row,
            offset_end_row=offset_end_row,
            offset_start_col=offset_start_col,
            offset_end_col=offset_end_col,
            clip_to_bounds=clip,
        )
    
    
    
    # ==================== 表头处理方法 ====================
    def apply_header(
        self,
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
                    names = _generate_placeholder_names(len(self.columns))
                elif len(names) != len(self.columns):
                    raise ValueError(f"提供的列名数量为 {len(names)}，与 DataFrame 列数 {len(self.columns)} 不一致")
            cleaned = [_safe_name(x) for x in names]
            cleaned = _dedup_names(cleaned)
            if inplace:
                self.columns = cleaned
                self.reset_index(drop=True, inplace=True)
                return None
            else:
                out = self.copy()
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
                placeholder_names = _generate_placeholder_names(len(self.columns))
                if inplace:
                    self.columns = placeholder_names
                    self.reset_index(drop=True, inplace=True)
                    return None
                else:
                    out = self.copy()
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
                    merged = header_join.join(map(str, unique_ordered)) if unique_ordered else str(self.columns[col_idx])
                    new_headers.append(_safe_name(merged))
                new_columns = _dedup_names(new_headers)
            
            if inplace:
                self.columns = new_columns
                self.reset_index(drop=True, inplace=True)
                return None
            else:
                out = self.copy()
                out.columns = new_columns
                out.reset_index(drop=True, inplace=True)
                return out

        # 3) header 为 False：不处理
        if header is False:
            if inplace:
                return None
            else:
                return self.copy()

        # 4) 与 pandas read_csv 语义对齐：
        #    - header=True 等价于 header=0（使用第0行做表头）
        #    - header=int 使用该“0基”行作为表头
        #    - header=[i,j] 使用多行表头
        if header is True:
            header = 0

        # 多行表头：list[int]
        if isinstance(header, (list, tuple)) and len(header) > 0 and isinstance(next(iter(header)), (int, np.integer)):
            idxs = list(header)
            if min(idxs) < 0 or max(idxs) >= len(self):
                raise ValueError("header 行索引超出范围")
            header_block = self.iloc[idxs, :].astype("string")
            data_start = max(idxs) + 1
            data_block = self.iloc[data_start:, :].copy()
            arrays = [header_block.iloc[i - idxs[0]].tolist() for i in idxs]
            if header_join is None:
                tuples = [tuple(items) for items in zip(*arrays)]
                data_block.columns = pd.MultiIndex.from_tuples(tuples)
            else:
                new_headers: List[str] = []
                for col_idx, items in enumerate(zip(*arrays)):
                    values = [x for x in items if pd.notna(x)]
                    unique_ordered = list(dict.fromkeys(values))
                    merged = header_join.join(map(str, unique_ordered)) if unique_ordered else str(self.columns[col_idx])
                    new_headers.append(_safe_name(merged))
                data_block.columns = _dedup_names(new_headers)
            data_block.reset_index(drop=True, inplace=True)
            if inplace:
                # 对于多行表头，需要替换整个 DataFrame
                self.__init__(data_block)
                return None
            else:
                return data_block

        # 单行表头：int
        if isinstance(header, (int, np.integer)):
            row_idx = int(header)
            if row_idx < 0 or row_idx >= len(self):
                raise ValueError("header 行索引超出范围")
            hdr = self.iloc[row_idx, :].astype("string")
            data_block = self.iloc[row_idx + 1:, :].copy()
            new_cols = [ _safe_name(x) for x in hdr.tolist() ]
            data_block.columns = _dedup_names(new_cols)
            data_block.reset_index(drop=True, inplace=True)
            if inplace:
                # 对于单行表头，需要替换整个 DataFrame
                self.__init__(data_block)
                return None
            else:
                return data_block

        # 其余类型不支持
        raise ValueError("不支持的 header 类型。支持: False/True/int/list[int]/list[str]/Series/DataFrame")

        
    def find_idx(
        self,
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
            if 0 <= target < len(self.columns):
                column_name = self.columns[target]
                column_data = self[column_name]
                if not isinstance(column_data, XlSeries):
                    column_data = XlSeries(column_data)
                return column_data.find_idx(q, mode=mode, na=na, flags=flags, nth=nth)
            else:
                # 列索引越界时回退为按行搜索
                axis = "row"
        
        if axis == "column":
            # 按列搜索
            if target not in self.columns:
                raise ValueError(f"列 '{target}' 不存在")
            column_data = self[target]
            if not isinstance(column_data, XlSeries):
                column_data = XlSeries(column_data)
            return column_data.find_idx(q, mode=mode, na=na, flags=flags, nth=nth)
        
        elif axis == "row":
            # 按行搜索 - 复用Series的find_idx方法
            if target not in self.index:
                raise ValueError(f"行索引 '{target}' 不存在")
            
            # 获取指定行的数据，并确保是XlSeries类型
            row_data = self.loc[target]
            if not isinstance(row_data, XlSeries):
                row_data = XlSeries(row_data)
            
            # 直接调用Series的find_idx方法
            return row_data.find_idx(q, mode=mode, na=na, flags=flags, nth=nth)
        
        else:
            raise ValueError("axis must be 'column' or 'row'")
    
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
    
    # ==================== Series 扩展方法 ====================
    
    def find_idx(
        self,
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
            arr = self.to_numpy(copy=False)
            if pd.isna(q):
                # 查找缺失值位置（覆盖 None/np.nan/pd.NA 等）
                idx = np.flatnonzero(pd.isna(arr))
            else:
                idx = np.flatnonzero(arr == q)
        elif mode == "contains":
            # contains：字面子串匹配（regex=False），避免正则引擎开销与语义歧义
            arr = self.astype("string")
            mask = arr.str.contains(str(q), regex=False, na=na)
            idx = np.flatnonzero(mask.to_numpy())
        elif mode == "regex":
            # regex：正则匹配，可通过 flags 控制大小写等
            arr = self.astype("string")
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