import os, sys
# 确保项目根目录在 sys.path 中
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)
    
import pandas as pd
from xlgrab.reader import get_sheet, get_region, get_cell, last_data_row


def test_reader():
    """简单测试reader功能"""
    file_path = "D:/data/demo.xlsx"
    sheet = "Sheet1"
    header_range = "A2:B2"
    
    try:
        # 测试get_sheet
        df = get_sheet(file_path, sheet)
        if df is not None:
            print(f"get_sheet成功: {df.shape}")
        else:
            print("get_sheet失败: 文件或sheet不存在")
            return
        
        # 测试get_region - 获取header区域
        region_df = get_region(file_path, sheet, 2, 2, 1, 2)  # A2:B2
        if region_df is not None:
            print(f"get_region成功: {region_df.shape}")
            print(f"Header内容: {region_df.iloc[0, 0]}, {region_df.iloc[0, 1]}")
        else:
            print("get_region失败")
        
        # 测试get_cell - 获取A2单元格
        cell_value = get_cell(file_path, sheet, 2, 1)  # A2
        print(f"get_cell A2: {cell_value}")
        
        # 测试get_cell - 获取B2单元格
        cell_value = get_cell(file_path, sheet, 2, 2)  # B2
        print(f"get_cell B2: {cell_value}")
        
        # 测试last_data_row - 获取A列最后数据行
        last_row = last_data_row(file_path, sheet, 1)  # A列
        print(f"last_data_row A列: {last_row}")
        
        # 测试last_data_row - 获取B列最后数据行
        last_row = last_data_row(file_path, sheet, 2)  # B列
        print(f"last_data_row B列: {last_row}")
        
        print("所有测试完成！")
        
    except Exception as e:
        print(f"测试出错: {e}")


if __name__ == "__main__":
    test_reader()