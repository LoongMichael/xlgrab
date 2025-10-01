from typing import Tuple
from openpyxl.utils import column_index_from_string, coordinate_to_tuple

"""
通用工具（utils）：目前仅包含坐标相关函数。
"""


def col_to_index(col: str) -> int:
    return int(column_index_from_string(col))


def a1_to_row_col(a1: str) -> Tuple[int, int]:
    row, col = coordinate_to_tuple(a1)
    return int(row), int(col)


