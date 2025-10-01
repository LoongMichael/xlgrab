"""
xlgrab 基本使用示例

展示极简API的各种用法
"""

import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from xlgrab import (
    extract_simple, 
    extract_with_header, 
    extract_table, 
    extract_list,
    range_spec,
    anchor_spec,
    extract
)


def example_1_simple_extraction():
    """示例1: 简单区域提取"""
    print("=== 示例1: 简单区域提取 ===")
    
    # 提取A1:C10区域
    result = extract_simple("data.xlsx", "Sheet1", "A1:C10")
    
    if result.data:
        print(f"提取到 {len(result.data)} 行数据")
        for i, row in enumerate(result.data[:3]):  # 显示前3行
            print(f"第{i+1}行: {row}")
    else:
        print("提取失败:", result.errors)


def example_2_header_extraction():
    """示例2: 带表头提取"""
    print("\n=== 示例2: 带表头提取 ===")
    
    # 提取表头和数据
    result = extract_with_header("data.xlsx", "Sheet1", "A1:C1", "A2:C10")
    
    if result.data:
        print(f"列名: {result.columns}")
        print(f"数据: {result.data[:2]}")  # 显示前2行数据
        
        # 转换为DataFrame
        df = result.to_dataframe()
        print(f"\nDataFrame:\n{df.head()}")
        
        # 转换为字典列表
        dict_list = result.to_dict()
        print(f"\n字典列表: {dict_list[:2]}")
    else:
        print("提取失败:", result.errors)


def example_3_table_extraction():
    """示例3: 自动表格提取"""
    print("\n=== 示例3: 自动表格提取 ===")
    
    # 自动检测表格边界
    result = extract_table("data.xlsx", "Sheet1", "A1")
    
    if result.data:
        print(f"自动检测到 {len(result.data)} 行数据")
        print(f"数据: {result.data[:2]}")
    else:
        print("提取失败:", result.errors)


def example_4_list_extraction():
    """示例4: 列表数据提取"""
    print("\n=== 示例4: 列表数据提取 ===")
    
    # 提取A列列表
    result = extract_list("data.xlsx", "Sheet1", "A", 2)  # 从第2行开始
    
    if result.data:
        print(f"列表数据: {[row[0] for row in result.data]}")
    else:
        print("提取失败:", result.errors)


def example_5_anchor_extraction():
    """示例5: 锚点提取"""
    print("\n=== 示例5: 锚点提取 ===")
    
    # 通过文本查找区域
    specs = [
        anchor_spec("Sheet1", "A", "姓名", 1, (1, 0))  # 在A列找"姓名"，向下偏移1行
    ]
    result = extract("data.xlsx", specs)
    
    if result.data:
        print(f"锚点数据: {result.data}")
    else:
        print("提取失败:", result.errors)


def example_6_multiple_regions():
    """示例6: 多区域提取"""
    print("\n=== 示例6: 多区域提取 ===")
    
    # 提取多个区域
    specs = [
        range_spec("Sheet1", "A1:C5"),      # 第一个区域
        range_spec("Sheet1", "A7:C10"),     # 第二个区域
        anchor_spec("Sheet1", "A", "总计", 1, (1, 0))  # 锚点区域
    ]
    result = extract("data.xlsx", specs)
    
    if result.data:
        print(f"多区域数据: {len(result.data)} 行")
        print(f"数据: {result.data[:3]}")
    else:
        print("提取失败:", result.errors)


def example_7_error_handling():
    """示例7: 错误处理"""
    print("\n=== 示例7: 错误处理 ===")
    
    # 尝试提取不存在的文件
    result = extract_simple("nonexistent.xlsx", "Sheet1", "A1:C10")
    
    if result.errors:
        print("错误信息:")
        for error in result.errors:
            print(f"  - {error}")
    else:
        print("意外成功:", result.data)


def run_all_examples():
    """运行所有示例"""
    print("xlgrab 使用示例")
    print("=" * 50)
    
    example_1_simple_extraction()
    example_2_header_extraction()
    example_3_table_extraction()
    example_4_list_extraction()
    example_5_anchor_extraction()
    example_6_multiple_regions()
    example_7_error_handling()
    
    print("\n所有示例完成！")


if __name__ == "__main__":
    run_all_examples()
