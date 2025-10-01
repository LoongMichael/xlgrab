"""
测试核心功能 - 验证极简API设计
"""

import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import pandas as pd
import tempfile
from xlgrab import extract_simple, extract_with_header, extract_table, extract_list, range_spec, anchor_spec, extract


def create_test_excel():
    """创建测试Excel文件"""
    # 创建测试数据
    data = {
        'A': ['姓名', 'Alice', 'Bob', 'Charlie', ''],
        'B': ['年龄', 25, 30, 35, ''],
        'C': ['城市', 'New York', 'London', 'Tokyo', ''],
        'D': ['部门', 'IT', 'HR', 'Finance', ''],
    }
    df = pd.DataFrame(data)
    
    # 创建临时Excel文件
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        df.to_excel(tmp.name, index=False, header=False, engine='openpyxl')
        return tmp.name


def test_extract_simple():
    """测试简单提取"""
    print("=== 测试 extract_simple ===")
    
    file_path = create_test_excel()
    try:
        # 提取固定区域
        result = extract_simple(file_path, "Sheet1", "A1:C3")
        print(f"数据形状: {len(result.data)} x {len(result.data[0]) if result.data else 0}")
        print(f"数据: {result.data}")
        print(f"错误: {result.errors}")
        print()
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass


def test_extract_with_header():
    """测试带表头提取"""
    print("=== 测试 extract_with_header ===")
    
    file_path = create_test_excel()
    try:
        # 提取带表头的数据
        result = extract_with_header(file_path, "Sheet1", "A1:C1", "A2:C4")
        print(f"列名: {result.columns}")
        print(f"数据形状: {len(result.data)} x {len(result.data[0]) if result.data else 0}")
        print(f"数据: {result.data}")
        print(f"错误: {result.errors}")
        print()
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass


def test_extract_table():
    """测试表格提取"""
    print("=== 测试 extract_table ===")
    
    file_path = create_test_excel()
    try:
        # 自动检测表格边界
        result = extract_table(file_path, "Sheet1", "A1")
        print(f"列名: {result.columns}")
        print(f"数据形状: {len(result.data)} x {len(result.data[0]) if result.data else 0}")
        print(f"数据: {result.data}")
        print(f"错误: {result.errors}")
        print()
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass


def test_extract_list():
    """测试列表提取"""
    print("=== 测试 extract_list ===")
    
    file_path = create_test_excel()
    try:
        # 提取A列列表
        result = extract_list(file_path, "Sheet1", "A", 2)  # 从第2行开始
        print(f"列表数据: {result.data}")
        print(f"错误: {result.errors}")
        print()
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass


def test_anchor_extraction():
    """测试锚点提取"""
    print("=== 测试锚点提取 ===")
    
    file_path = create_test_excel()
    try:
        # 通过锚点查找
        specs = [
            anchor_spec("Sheet1", "A", "姓名", 1, (1, 0))  # 在A列找"姓名"，向下偏移1行
        ]
        result = extract(file_path, specs)
        print(f"锚点数据: {result.data}")
        print(f"错误: {result.errors}")
        print()
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass


def test_dataframe_conversion():
    """测试DataFrame转换"""
    print("=== 测试DataFrame转换 ===")
    
    file_path = create_test_excel()
    try:
        result = extract_with_header(file_path, "Sheet1", "A1:C1", "A2:C4")
        df = result.to_dataframe()
        print(f"DataFrame形状: {df.shape}")
        print(f"DataFrame:\n{df}")
        print()
        
        # 测试字典转换
        dict_list = result.to_dict()
        print(f"字典列表: {dict_list}")
        print()
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass


def run_all_tests():
    """运行所有测试"""
    print("开始测试极简API设计")
    print("=" * 50)
    
    test_extract_simple()
    test_extract_with_header()
    test_extract_table()
    test_extract_list()
    test_anchor_extraction()
    test_dataframe_conversion()
    
    print("所有测试完成！")


if __name__ == "__main__":
    run_all_tests()
