import os, sys

# 确保项目根目录在 sys.path 中
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from xlgrab.reader import WorkbookReader
from xlgrab.locator import locator_regions
from xlgrab.models import Rule, HeaderSpec, BlockSpec, AnchorSpec, EndSpec
from xlgrab.utils import a1_to_row_col


def width_from_header(header_range: str) -> int:
    hs, he = header_range.split(":")
    _, sc = a1_to_row_col(hs)
    _, ec = a1_to_row_col(he)
    return ec - sc + 1


def test_anchor_then_keyword_end():
    # 场景描述：
    # 文件: D:/data/anchor.xlsx
    # sheet: Sheet1
    # header: A2:B2
    # 锚点：在 A 列查找“名称”的第 3 次出现；数据从锚点下一行开始（row_delta=+1）
    # 终止关键词: “合计” 的上一行（offset=-1）
    # total: 无
    file_path = "D:/data/anchor.xlsx"   # 请替换为实际文件路径
    sheet = "Sheet1"
    header_range = "A2:B2"

    rule = Rule(
        rule_id="anchor_kw",
        sheet_name=sheet,
        header=HeaderSpec(header_range=header_range),
        blocks=[
            BlockSpec(
                type="anchor",
                start=AnchorSpec(in_column="A", find_text="名称", occurrence=3, offset=(1, 0)),  # 下一行
                end=EndSpec(mode="by_keyword_with_offset", keyword="合计", offset_rows=-1),
                width=width_from_header(header_range),
                include_header=False,
            )
        ],
        clean_empty_rows=False,
    )

    reader = WorkbookReader(file_path=file_path)
    regions, errors = locator_regions(reader, [rule])

    print("regions:", [
        {
            "sheet": r.sheet_name,
            "data": {
                "start_row": r.data_start_row,
                "end_row": r.data_end_row,
                "start_col": r.data_start_col,
                "end_col": r.data_start_col + r.width - 1,
            }
        } for r in regions
    ])
    print("errors:", [(e.reason_code, e.message) for e in errors])


if __name__ == "__main__":
    test_anchor_then_keyword_end()


