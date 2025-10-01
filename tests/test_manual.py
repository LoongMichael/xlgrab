import os
import sys
import json

# 允许从项目根目录导入包
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from xlgrab.locator import locate_batch


def main() -> None:
    file_path = r"D:\data\demo.xlsx"
    sheet_name = "Sheet1"

    print("=== xlgrab demo: fixed ranges ===")
    print(f"file: {file_path}")
    print(f"sheet: {sheet_name}")

    if not os.path.exists(file_path):
        print(f"文件不存在，跳过测试: {file_path}")
        return

    operations = [
        {"type": "rows_by_range", "name": "rows", "params": {"area": "A2:B9"}},
        {"type": "columns_by_range", "name": "cols", "params": {"area": "A2:B9"}},
        {"type": "region_by_range", "name": "region", "params": {"area": "A2:B9"}},
        {
            "type": "regions_by_range",
            "name": "regions",
            "params": {
                "items": [
                    {"name": "r1", "area": "A2:B9"},
                    {"name": "r2", "area": "C3:D5"},
                ]
            },
        },
        # 数字/字符串列行示例（可能因关键词不存在而返回 None，仅用于演示传参形式）
        {
            "type": "rows_by_keywords",
            "name": "rows_kw_numeric_col",
            "params": {
                "start_col": 1,               # 列可用数字 1 表示 'A'
                "start_keyword": "开始",
                "end_col": "2",             # 列也可用字符串数字 "2" 表示第2列
                "end_keyword": "结束"
            },
        },
        {
            "type": "columns_by_keywords",
            "name": "cols_kw_numeric_row",
            "params": {
                "header_row": "1",          # 行可用字符串数字 "1"
                "start_keyword": "姓名",
                "end_keyword": "工资",
                "contains": False
            },
        },
        {
            "type": "region_by_specs",
            "name": "spec_region",
            "params": {
                "row": {"mode": "range", "area": "A2:B9"},
                "col": {"mode": "range", "area": "A2:B9"},
            },
        },
        {
            "type": "regions_by_specs",
            "name": "spec_regions",
            "params": {
                "items": [
                    {
                        "name": "r1",
                        "row": {"mode": "range", "area": "A2:B9"},
                        "col": {"mode": "range", "area": "A2:B9"},
                        "offsets": {"end_row": -2}  # 表尾上移两行示例
                    },
                    {
                        "name": "r2",
                        "row": {"mode": "range", "area": "C3:D5"},
                        "col": {"mode": "range", "area": "C3:D5"},
                    },
                ]
            ,
                "offsets": {"start_col": 1}  # 批量默认偏移，示例：起始列右移1
            },
        },
    ]

    results = locate_batch(file_path, sheet_name, operations)
    print("\n结果:")
    print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()


