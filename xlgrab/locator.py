"""
区域定位器 - 高性能批量定位架构

支持6种定位方法，通过批量接口一次读取，多次定位：
行定位：直接区域、关键词定位、关键词+最后行
列定位：直接区域、关键词定位、关键词+最后列
"""

from typing import Optional, Tuple, Union, Dict
import pandas as pd
from .reader import get_sheet
from .utils import a1_to_row_col, col_to_index


# ============================================================================
# 常量定义 - 操作类型（避免字符串硬编码散落）
# ============================================================================
OP_ROWS_BY_RANGE = 'rows_by_range'
OP_ROWS_BY_KEYWORDS = 'rows_by_keywords'
OP_ROWS_BY_START_KEYWORD = 'rows_by_start_keyword'
OP_COLUMNS_BY_RANGE = 'columns_by_range'
OP_COLUMNS_BY_KEYWORDS = 'columns_by_keywords'
OP_COLUMNS_BY_START_KEYWORD = 'columns_by_start_keyword'
OP_REGION_BY_RANGE = 'region_by_range'
OP_REGIONS_BY_RANGE = 'regions_by_range'
OP_REGION_BY_SPECS = 'region_by_specs'
OP_REGIONS_BY_SPECS = 'regions_by_specs'


# ============================================================================
# 批量定位接口 - 高性能定位解决方案
# ----------------------------------------------------------------------------
# 单一入口：一次读取 `sheet` 后执行多种定位操作。
# 支持的 op_type：
#   - 行：rows_by_range | rows_by_keywords | rows_by_start_keyword
#   - 列：columns_by_range | columns_by_keywords | columns_by_start_keyword
#   - 区域：region_by_range | regions_by_range（复用行/列定位）
# ============================================================================

def locate_batch(file_path: str, sheet: str, operations: list) -> dict:
    """批量定位接口，一次读取sheet，执行多个定位操作
    
    Args:
        file_path: Excel文件路径
        sheet: 工作表名称
        operations: 定位操作列表，每个操作包含type和参数
    
    Returns:
        dict: 定位结果字典，key为操作名称，value为定位结果
    """
    try:
        # 一次性读取sheet
        df = get_sheet(file_path, sheet)
        if df is None or df.empty:
            return {}
        
        results = {}

        # [EDIT] 使用分发表简化分支逻辑，便于扩展与维护
        dispatch = {
            OP_ROWS_BY_RANGE: _locate_rows_by_range_internal,
            OP_ROWS_BY_KEYWORDS: _locate_rows_by_keywords_internal,
            OP_ROWS_BY_START_KEYWORD: _locate_rows_by_start_keyword_internal,
            OP_COLUMNS_BY_RANGE: _locate_columns_by_range_internal,
            OP_COLUMNS_BY_KEYWORDS: _locate_columns_by_keywords_internal,
            OP_COLUMNS_BY_START_KEYWORD: _locate_columns_by_start_keyword_internal,
            OP_REGION_BY_RANGE: _locate_region_by_range_internal,
            OP_REGIONS_BY_RANGE: _locate_regions_by_range_internal,
            OP_REGION_BY_SPECS: _locate_region_by_specs_internal,
            OP_REGIONS_BY_SPECS: _locate_regions_by_specs_internal,
        }

        for op in operations:
            op_type = op.get('type')
            op_name = op.get('name', f"op_{len(results)}")
            op_params = op.get('params', {})

            handler = dispatch.get(op_type)
            result = handler(df, op_params) if handler else None

            results[op_name] = result
        
        return results
        
    except Exception:
        return {}


# ============================================================================
# 内部函数 - 直接操作DataFrame，用于批量定位
# ----------------------------------------------------------------------------
# 分组说明：
#   1) 行定位
#   2) 列定位
#   3) 区域定位（基于行/列定位复用组合）
#   4) 基础查找（关键词行/列）
# ============================================================================

def _is_valid_area(area: Optional[str]) -> bool:
    """校验 A1 区域字符串的基本合法性（语法与坐标顺序）"""
    if not area or ":" not in area:
        return False
    try:
        start_cell, end_cell = area.split(":", 1)
        start_row, start_col = a1_to_row_col(start_cell)
        end_row, end_col = a1_to_row_col(end_cell)
        if start_row <= 0 or start_col <= 0 or end_row <= 0 or end_col <= 0:
            return False
        if start_row > end_row or start_col > end_col:
            return False
        return True
    except Exception:
        return False


def _normalize_column(column: Union[str, int]) -> Optional[int]:
    """将列规范标准化为 1-based 列号，支持 'A' 或 1（字符串数字也可）。"""
    try:
        # int 直接返回（需 >0）
        if isinstance(column, int):
            return column if column > 0 else None
        if not isinstance(column, str):
            return None
        s = column.strip()
        if not s:
            return None
        # 纯数字字符串
        if s.isdigit():
            val = int(s)
            return val if val > 0 else None
        # 列字母，如 'A', 'BC'
        return col_to_index(s)
    except Exception:
        return None


def _normalize_row(row: Union[str, int]) -> Optional[int]:
    """将行规范标准化为 1-based 行号，支持 '1' 或 1。"""
    try:
        if isinstance(row, int):
            return row if row > 0 else None
        if not isinstance(row, str):
            return None
        s = row.strip()
        if not s.isdigit():
            return None
        val = int(s)
        return val if val > 0 else None
    except Exception:
        return None


def _locate_rows_by_range_internal(df: pd.DataFrame, params: dict) -> Optional[Tuple[int, int]]:
    """内部函数：直接指定区域定位行"""
    try:
        area = params.get('area')
        if not _is_valid_area(area):
            return None
        
        start_cell, end_cell = area.split(":", 1)
        start_row, _ = a1_to_row_col(start_cell)
        end_row, _ = a1_to_row_col(end_cell)
        
        if start_row <= 0 or end_row <= 0 or start_row > end_row:
            return None
        
        return (start_row, end_row)
    except Exception:
        return None


 


def _locate_rows_by_keywords_internal(df: pd.DataFrame, params: dict) -> Optional[Tuple[int, int]]:
    """内部函数：关键词定位开始和结束行

    [EDIT] 扩展参数：
      - contains: 是否子串匹配（默认 False）
      - occurrence_start / occurrence_end: 第几个匹配（默认 1；-1 表示取最后一个）
      - regex: contains 下是否启用正则（默认 False）
      - case_sensitive: 是否区分大小写（默认 True）
    """
    try:
        start_col = params.get('start_col')
        start_keyword = params.get('start_keyword')
        end_col = params.get('end_col')
        end_keyword = params.get('end_keyword')
        contains = params.get('contains', False)  # [EDIT] 新增
        regex = params.get('regex', False)  # [EDIT] 新增
        case_sensitive = params.get('case_sensitive', True)  # [EDIT] 新增
        occurrence_start = params.get('occurrence_start', 1)  # [EDIT] 新增
        occurrence_end = params.get('occurrence_end', 1)  # [EDIT] 新增
        
        if not all([start_col, start_keyword, end_col, end_keyword]):
            return None
        
        start_row = _find_keyword_row_internal(
            df,
            start_col,
            start_keyword,
            occurrence=occurrence_start,
            contains=contains,
            regex=regex,
            case_sensitive=case_sensitive,
        )  # [EDIT] 传递新增参数
        if start_row is None:
            return None
        
        end_row = _find_keyword_row_internal(
            df,
            end_col,
            end_keyword,
            occurrence=occurrence_end,
            contains=contains,
            regex=regex,
            case_sensitive=case_sensitive,
        )  # [EDIT] 传递新增参数
        if end_row is None:
            return None
        
        if start_row > end_row:
            return None
        
        return (start_row, end_row)
    except Exception:
        return None


def _locate_rows_by_start_keyword_internal(df: pd.DataFrame, params: dict) -> Optional[Tuple[int, int]]:
    """内部函数：关键词定位开始行，结束行使用最后数据行

    [EDIT] 扩展参数：
      - contains, occurrence, regex, case_sensitive 同上
    """
    try:
        start_col = params.get('start_col')
        start_keyword = params.get('start_keyword')
        contains = params.get('contains', False)  # [EDIT] 新增
        regex = params.get('regex', False)  # [EDIT] 新增
        case_sensitive = params.get('case_sensitive', True)  # [EDIT] 新增
        occurrence = params.get('occurrence', 1)  # [EDIT] 新增
        
        if not start_col or not start_keyword:
            return None
        
        start_row = _find_keyword_row_internal(
            df,
            start_col,
            start_keyword,
            occurrence=occurrence,
            contains=contains,
            regex=regex,
            case_sensitive=case_sensitive,
        )  # [EDIT] 传递新增参数
        if start_row is None:
            return None
        
        end_row = int(df.shape[0])
        if end_row == 0 or end_row < start_row:
            return None
        
        return (start_row, end_row)
    except Exception:
        return None


def _locate_columns_by_range_internal(df: pd.DataFrame, params: dict) -> Optional[Tuple[int, int]]:
    """内部函数：直接指定区域定位列范围"""
    try:
        area = params.get('area')
        if not _is_valid_area(area):
            return None
        
        start_cell, end_cell = area.split(":", 1)
        _, start_col = a1_to_row_col(start_cell)
        _, end_col = a1_to_row_col(end_cell)
        
        if start_col <= 0 or end_col <= 0 or start_col > end_col:
            return None
        
        return (start_col, end_col)
    except Exception:
        return None


def _locate_columns_by_keywords_internal(df: pd.DataFrame, params: dict) -> Optional[Tuple[int, int]]:
    """内部函数：关键词定位开始和结束列

    [EDIT] 扩展参数：
      - regex: contains 下是否启用正则（默认 False）
      - case_sensitive: 是否区分大小写（默认 True）
      - occurrence_start / occurrence_end: 第几个匹配（默认 1；-1 取最后）
    """
    try:
        header_row = params.get('header_row')
        start_keyword = params.get('start_keyword')
        end_keyword = params.get('end_keyword')
        contains = params.get('contains', False)
        regex = params.get('regex', False)  # [EDIT] 新增
        case_sensitive = params.get('case_sensitive', True)  # [EDIT] 新增
        occurrence_start = params.get('occurrence_start', 1)  # [EDIT] 新增
        occurrence_end = params.get('occurrence_end', -1)  # [EDIT] 新增，默认从右向左找最后一个
        
        if not all([header_row, start_keyword, end_keyword]):
            return None
        
        start_col = _find_keyword_column_internal(
            df,
            header_row,
            start_keyword,
            occurrence=occurrence_start,
            contains=contains,
            regex=regex,
            case_sensitive=case_sensitive,
        )  # [EDIT]
        if start_col is None:
            return None
        
        end_col = _find_keyword_column_internal(
            df,
            header_row,
            end_keyword,
            occurrence=occurrence_end,
            contains=contains,
            regex=regex,
            case_sensitive=case_sensitive,
        )  # [EDIT]
        if end_col is None:
            return None
        
        if start_col > end_col:
            return None
        
        return (start_col, end_col)
    except Exception:
        return None


def _locate_columns_by_start_keyword_internal(df: pd.DataFrame, params: dict) -> Optional[Tuple[int, int]]:
    """内部函数：关键词定位开始列，结束列使用最后数据列

    [EDIT] 扩展参数：contains/regex/case_sensitive/occurrence
    """
    try:
        header_row = params.get('header_row')
        start_keyword = params.get('start_keyword')
        contains = params.get('contains', False)
        regex = params.get('regex', False)  # [EDIT] 新增
        case_sensitive = params.get('case_sensitive', True)  # [EDIT] 新增
        occurrence = params.get('occurrence', 1)  # [EDIT] 新增
        
        if not header_row or not start_keyword:
            return None
        
        start_col = _find_keyword_column_internal(
            df,
            header_row,
            start_keyword,
            occurrence=occurrence,
            contains=contains,
            regex=regex,
            case_sensitive=case_sensitive,
        )  # [EDIT]
        if start_col is None:
            return None
        
        end_col = int(df.shape[1])
        if start_col > end_col:
            return None
        
        return (start_col, end_col)
    except Exception:
        return None


# ============================================================================
# 区域定位（基于行/列定位复用组合）
# ============================================================================
def _locate_region_by_range_internal(df: pd.DataFrame, params: dict) -> Optional[Tuple[int, int, int, int]]:
    """内部函数：使用现有行/列定位实现单区域四元组输出"""
    try:
        area = params.get('area')
        if not _is_valid_area(area):
            return None
        rows = _locate_rows_by_range_internal(df, {'area': area})
        cols = _locate_columns_by_range_internal(df, {'area': area})
        if rows is None or cols is None:
            return None
        start_row, end_row = rows
        start_col, end_col = cols
        # 偏移支持（可选）：params.offsets = {start_row,end_row,start_col,end_col}，允许运行负值
        offsets = params.get('offsets') or {}
        if offsets:
            adjusted = _apply_region_offsets((start_row, end_row, start_col, end_col), offsets, df)
            return adjusted
        return (start_row, end_row, start_col, end_col)
    except Exception:
        return None


def _locate_regions_by_range_internal(df: pd.DataFrame, params: dict) -> Optional[dict]:
    """内部函数：批量固定区域四元组输出，复用行/列定位方法"""
    try:
        items = params.get('items')
        if not isinstance(items, list):
            return None
        out: dict = {}
        default_offsets = params.get('offsets') or {}
        for idx, item in enumerate(items):
            name = item.get('name', f'item_{idx}')
            area = item.get('area')
            offsets = item.get('offsets') or default_offsets
            region = _locate_region_by_range_internal(df, {'area': area, 'offsets': offsets})
            if region is not None:
                out[name] = region
        return out
    except Exception:
        return None


def _locate_region_by_specs_internal(df: pd.DataFrame, params: dict) -> Optional[Tuple[int, int, int, int]]:
    """统一区域定位：行/列各自支持三种模式的任意组合

    params 示例：
      {
        'row': { 'mode': 'range'|'keywords'|'start_keyword', ...对应参数... },
        'col': { 'mode': 'range'|'keywords'|'start_keyword', ...对应参数... }
      }
    行各模式所需参数：
      - range: {'area': 'A2:B9'}
      - keywords: {'start_col': 'A', 'start_keyword': '开始', 'end_col': 'B', 'end_keyword': '结束'}
      - start_keyword: {'start_col': 'A', 'start_keyword': '开始'}
    列各模式所需参数：
      - range: {'area': 'A2:B9'}
      - keywords: {'header_row': 1, 'start_keyword': '姓名', 'end_keyword': '工资', 'contains': False}
      - start_keyword: {'header_row': 1, 'start_keyword': '姓名', 'contains': False}
    """
    try:
        row_spec = params.get('row') or {}
        col_spec = params.get('col') or {}

        row_mode = (row_spec.get('mode') or '').strip()
        col_mode = (col_spec.get('mode') or '').strip()
        if not row_mode or not col_mode:
            return None

        # 行定位分发
        if row_mode == 'range':
            rows = _locate_rows_by_range_internal(df, row_spec)
        elif row_mode == 'keywords':
            rows = _locate_rows_by_keywords_internal(df, row_spec)
        elif row_mode == 'start_keyword':
            rows = _locate_rows_by_start_keyword_internal(df, row_spec)
        else:
            return None

        if rows is None:
            return None

        # 列定位分发
        if col_mode == 'range':
            cols = _locate_columns_by_range_internal(df, col_spec)
        elif col_mode == 'keywords':
            cols = _locate_columns_by_keywords_internal(df, col_spec)
        elif col_mode == 'start_keyword':
            cols = _locate_columns_by_start_keyword_internal(df, col_spec)
        else:
            return None

        if cols is None:
            return None

        start_row, end_row = rows
        start_col, end_col = cols
        region = (start_row, end_row, start_col, end_col)
        offsets = params.get('offsets') or {}
        if offsets:
            return _apply_region_offsets(region, offsets, df)
        return region
    except Exception:
        return None


def _locate_regions_by_specs_internal(df: pd.DataFrame, params: dict) -> Optional[dict]:
    """批量统一区域定位：支持 items 内部任意行/列模式组合

    params 示例：
      {
        'items': [
           { 'name': 'r1', 'row': {...}, 'col': {...} },
           { 'name': 'r2', 'row': {...}, 'col': {...} },
        ]
      }
    返回：{ name: (r1, r2, c1, c2), ... }
    """
    try:
        items = params.get('items')
        if not isinstance(items, list):
            return None
        out: dict = {}
        default_offsets = params.get('offsets') or {}
        for idx, item in enumerate(items):
            name = item.get('name', f'item_{idx}')
            offsets = item.get('offsets') or default_offsets
            region = _locate_region_by_specs_internal(
                df,
                {
                    'row': item.get('row') or {},
                    'col': item.get('col') or {},
                    'offsets': offsets,
                },
            )
            if region is not None:
                out[name] = region
        return out
    except Exception:
        return None


# ============================================================================
# 基础查找
# ============================================================================

def _apply_region_offsets(region: Tuple[int, int, int, int], offsets: dict, df: pd.DataFrame) -> Optional[Tuple[int, int, int, int]]:
    """对区域四元组应用偏移，并裁剪到有效范围。

    支持的偏移键：
      - 'start_row' | 'r1'
      - 'end_row'   | 'r2'
      - 'start_col' | 'c1'
      - 'end_col'   | 'c2'
    """
    try:
        sr, er, sc, ec = region
        def _get(key_primary: str, key_alt: str) -> int:
            val = offsets.get(key_primary, offsets.get(key_alt, 0))
            try:
                return int(val)
            except Exception:
                return 0

        sr += _get('start_row', 'r1')
        er += _get('end_row', 'r2')
        sc += _get('start_col', 'c1')
        ec += _get('end_col', 'c2')

        # 裁剪到有效范围（1-based）
        max_r = int(df.shape[0]) if df is not None else max(sr, er)
        max_c = int(df.shape[1]) if df is not None else max(sc, ec)
        sr = max(1, min(sr, max_r))
        er = max(1, min(er, max_r))
        sc = max(1, min(sc, max_c))
        ec = max(1, min(ec, max_c))

        if sr > er or sc > ec:
            return None
        return (sr, er, sc, ec)
    except Exception:
        return None
def _find_keyword_row_internal(df: pd.DataFrame, column: Union[str, int], text: str, 
                              occurrence: int = 1, contains: bool = False,
                              regex: bool = False, case_sensitive: bool = True) -> Optional[int]:
    """内部函数：在指定列中查找关键词的行号

    [EDIT] 新增参数：regex、case_sensitive，用于控制 contains 行为与大小写
    """
    try:
        col_1based = _normalize_column(column)
        if col_1based is None:
            return None
        col_idx = col_1based - 1
        if col_idx >= len(df.columns):
            return None
            
        series = df.iloc[:, col_idx].astype(str)

        # [EDIT] 大小写控制：非区分大小写时统一小写
        if not case_sensitive:
            series = series.str.lower()
            text_cmp = text.lower()
        else:
            text_cmp = text

        if contains:
            # [EDIT] contains 默认 regex=False 更安全，可由参数控制
            mask = series.str.contains(text_cmp, na=False, regex=regex)
        else:
            mask = series.str.strip() == text_cmp
            
        rows = mask[mask].index.tolist()
        
        if not rows:
            return None

        if occurrence == -1:
            return rows[-1] + 1

        if len(rows) >= occurrence:
            return rows[occurrence - 1] + 1
        
        return None
    except Exception:
        return None


def _find_keyword_column_internal(df: pd.DataFrame, row: Union[int, str], text: str, 
                                 occurrence: int = 1, contains: bool = False,
                                 regex: bool = False, case_sensitive: bool = True) -> Optional[int]:
    """内部函数：在指定行中查找关键词的列号

    [EDIT] 新增参数：regex、case_sensitive，用于控制 contains 行为与大小写
    """
    try:
        row_1based = _normalize_row(row)
        if row_1based is None:
            return None
        row_idx = row_1based - 1
        if row_idx < 0 or row_idx >= df.shape[0]:
            return None
            
        row_values = df.iloc[row_idx, :].astype(str)

        # [EDIT] 大小写控制
        if not case_sensitive:
            row_values = row_values.str.lower()
            text_cmp = text.lower()
        else:
            text_cmp = text

        row_values = row_values.str.strip()

        if contains:
            mask = row_values.str.contains(text_cmp, na=False, regex=regex)
        else:
            mask = row_values == text_cmp
            
        cols = mask[mask].index.tolist()
        
        if not cols:
            return None
        
        if occurrence == -1:
            return cols[-1] + 1
        
        if len(cols) >= occurrence:
            return cols[occurrence - 1] + 1
        
        return None
    except Exception:
        return None


 