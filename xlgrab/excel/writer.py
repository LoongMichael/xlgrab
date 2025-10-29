"""
Excel 写入模块

提供向现有Excel文件高效写入数据的功能。对外暴露三个核心函数：
- to_sheet_many: (推荐) 自动分批写入多个任务，性能最高。
- write_to_excel: 写入单个DataFrame。
- write_range_to_excel: 写入二维列表或元组。
"""

from itertools import groupby
from operator import itemgetter
from typing import Any, Dict, List, Literal, Optional, Union
import warnings

import openpyxl
import pandas as pd

# ====================================================================
# Public API
# ====================================================================

MergePolicy = Literal["unmerge", "error"]

def to_sheet_many(tasks: List[Dict[str, Any]]) -> None:
    """
    自动按文件名分批，向多个Excel文件高效写入数据。
    """
    sorted_tasks = sorted(tasks, key=itemgetter('excel_name'))
    for excel_name, group in groupby(sorted_tasks, key=itemgetter('excel_name')):
        with _ExcelBatchWriter(str(excel_name)) as writer:
            task_list = list(group)
            write_tasks = [{k: v for k, v in task.items() if k != 'excel_name'} for task in task_list]
            writer.write_many(write_tasks)

def write_to_excel(df: pd.DataFrame,
                   excel_name: str,
                   sheet_name: Union[str, int] = 0,
                   start_row: int = 1,
                   start_col: int = 1,
                   end_row: Optional[int] = None,
                   end_col: Optional[int] = None,
                   overwrite: bool = False,
                   header: bool = True,
                   index: bool = False,
                   merge_policy: MergePolicy = 'unmerge',
                   *,
                   _workbook: Optional[openpyxl.Workbook] = None,
                   _save: bool = True) -> None:
    """
    向现有Excel文件的指定位置写入DataFrame数据。
    """
    if not isinstance(df, pd.DataFrame):
        raise ValueError("df参数必须是pandas DataFrame")
    if not isinstance(excel_name, str):
        raise ValueError("excel_name参数必须是字符串")
    if start_row < 1 or start_col < 1:
        raise ValueError("start_row/start_col必须大于等于1")

    _perform_write(
        df, excel_name, sheet_name, start_row, start_col,
        end_row, end_col, overwrite, header, index, merge_policy, _workbook, _save
    )

def write_range_to_excel(data: Union[pd.DataFrame, list, tuple],
                         excel_name: str,
                         sheet_name: Union[str, int] = 0,
                         start_row: int = 1,
                         start_col: int = 1,
                         end_row: Optional[int] = None,
                         end_col: Optional[int] = None,
                         merge_policy: MergePolicy = 'unmerge') -> None:
    """向Excel文件的指定范围写入数据（简化版本）。"""
    if isinstance(data, pd.DataFrame):
        df = data
    elif isinstance(data, (list, tuple)):
        df = pd.DataFrame(data)
    else:
        raise ValueError("data参数必须是DataFrame、列表或元组")

    write_to_excel(
        df=df,
        excel_name=excel_name,
        sheet_name=sheet_name,
        start_row=start_row,
        start_col=start_col,
        end_row=end_row,
        end_col=end_col,
        overwrite=True,
        header=False,
        index=False,
        merge_policy=merge_policy,
    )

# ====================================================================
# Internal Implementation
# ====================================================================

class _ExcelBatchWriter:
    """内部类：一次打开、多次写、一次保存。"""
    def __init__(self, excel_name: str):
        self.excel_name = excel_name
        self.workbook = _open_or_create_workbook(excel_name)

    def write(self, **kwargs) -> None:
        write_to_excel(excel_name=self.excel_name, _workbook=self.workbook, _save=False, **kwargs)

    def write_many(self, tasks: List[Dict[str, Any]]) -> None:
        for task in tasks:
            self.write(**task)

    def save(self) -> None:
        self.workbook.save(self.excel_name)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.save()

def _perform_write(df: pd.DataFrame, excel_name: str, sheet_name: Union[str, int], start_row: int, start_col: int,
                   end_row: Optional[int], end_col: Optional[int], overwrite: bool, header: bool, index: bool,
                   merge_policy: MergePolicy, _workbook: Optional[openpyxl.Workbook], _save: bool) -> None:
    """包含所有写入逻辑的内部函数。"""
    # 1. 准备写入数据，并确定最终写入的行数和列数
    data_to_write = df.values.tolist()
    if header and not df.empty:
        col_names = [f"Column_{i}" for i in df.columns] if isinstance(df.columns, pd.RangeIndex) else df.columns.tolist()
        data_to_write.insert(0, col_names)
    if index and not df.empty:
        row_names = [f"Row_{i}" for i in df.index] if isinstance(df.index, pd.RangeIndex) else df.index.tolist()
        # 为 header 行的索引位置留空
        if header: row_names.insert(0, None) 
        for i, row in enumerate(data_to_write):
            data_to_write[i].insert(0, row_names[i])

    final_rows = len(data_to_write)
    final_cols = len(data_to_write[0]) if final_rows > 0 else 0

    # 2. 确定最终写入区域，并处理合并单元格
    write_end_row = start_row + final_rows - 1
    write_end_col = start_col + final_cols - 1

    wb = _workbook or _open_or_create_workbook(excel_name)
    ws = _get_or_create_worksheet(wb, sheet_name)
    _handle_merged_cells(ws, start_row, write_end_row, start_col, write_end_col, merge_policy)

    # 3. 写入数据 (注意：当前版本忽略了 end_row, end_col, overwrite 的截断逻辑，因为这与 unmerge 逻辑冲突)
    #    后续可以优化为先 unmerge，再根据截断后的尺寸写入
    cell_set = ws.cell
    for i, row in enumerate(data_to_write):
        r = start_row + i
        for j, val in enumerate(row):
            cell_set(row=r, column=start_col + j, value=val)

    if _workbook is None and _save: wb.save(excel_name)

def _handle_merged_cells(ws: openpyxl.worksheet.worksheet.Worksheet, start_row: int, end_row: int, start_col: int, end_col: int, policy: MergePolicy):
    """根据策略处理与写入区域重叠的合并单元格。"""
    if not (hasattr(ws, 'merged_cells') and ws.merged_cells.ranges):
        return

    # 必须复制，因为 unmerge_cells 会修改原始集合
    merged_ranges = list(ws.merged_cells.ranges)

    overlapping_ranges = []
    for m_range in merged_ranges:
        if not (m_range.max_row < start_row or m_range.min_row > end_row or
                m_range.max_col < start_col or m_range.min_col > end_col):
            overlapping_ranges.append(m_range)

    if not overlapping_ranges: return

    if policy == 'error':
        raise ValueError(f"写入区域与合并单元格 {overlapping_ranges[0]} 存在冲突。请更改写入位置或使用 merge_policy='unmerge'。")
    
    if policy == 'unmerge':
        for m_range in overlapping_ranges:
            # unmerge_cells 可能会在某些 openpyxl 版本中对已移除的 range 抛出 KeyError
            try:
                ws.unmerge_cells(str(m_range))
            except KeyError:
                pass

def _open_or_create_workbook(excel_name: str) -> openpyxl.Workbook:
    try: return openpyxl.load_workbook(excel_name, data_only=False, read_only=False)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        if 'Sheet' in wb.sheetnames: wb.remove(wb['Sheet'])
        return wb

def _get_or_create_worksheet(workbook: openpyxl.Workbook, sheet_name: Union[str, int]):
    if isinstance(sheet_name, int):
        if sheet_name < len(workbook.sheetnames): return workbook.worksheets[sheet_name]
        return workbook.create_sheet(f"Sheet{sheet_name + 1}")
    return workbook[sheet_name] if sheet_name in workbook.sheetnames else workbook.create_sheet(sheet_name)
