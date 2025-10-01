"""
测试固定输入路径与区域：D:\data\demo.xlsx, sheet1, A2:B9
"""

import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from xlgrab.locator import locate_batch


def test_demo_fixed_path():
    print("=== 测试固定路径与区域: D:\\data\\demo.xlsx / Sheet1 / A2:B9 ===")

    file_path = r"D:\data\demo.xlsx"
    sheet_name = "Sheet1"

    if not os.path.exists(file_path):
        print(f"文件不存在，跳过测试: {file_path}")
        return

    operations = [
        {'type': 'rows_by_range', 'name': 'rows', 'params': {'area': 'A2:B9'}},
        {'type': 'columns_by_range', 'name': 'cols', 'params': {'area': 'A2:B9'}},
    ]

    results = locate_batch(file_path, sheet_name, operations)
    print("结果:")
    print(f"  rows: {results.get('rows')}")
    print(f"  cols: {results.get('cols')}")
    if not results.get('rows') or not results.get('cols'):
        print("定位失败或sheet不存在，跳过断言")
        return


if __name__ == "__main__":
    test_demo_fixed_path()


