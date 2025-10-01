"""
测试列定位功能
"""

import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import pandas as pd
import tempfile
from xlgrab.locator import locate_batch


def create_test_excel():
    """创建测试Excel文件"""
    data = {
        'A': ['姓名', 'Alice', 'Bob', 'Charlie'],
        'B': ['年龄', 25, 30, 35],
        'C': ['城市', 'New York', 'London', 'Tokyo'],
        'D': ['部门', 'IT', 'HR', 'Finance'],
    }
    df = pd.DataFrame(data)
    
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        df.to_excel(tmp.name, index=False, header=False, engine='openpyxl')
        return tmp.name


def test_columns_by_range():
    """测试方法1: 直接指定列范围"""
    print("=== 测试列定位方法1: 直接指定列范围 ===")
    
    file_path = create_test_excel()
    try:
        operations = [
            {'type': 'columns_by_range', 'name': 'cols', 'params': {'area': 'A1:C1'}}
        ]
        results = locate_batch(file_path, "Sheet1", operations)
        if results.get('cols'):
            print(f"列范围定位成功: {results['cols']}")
        else:
            print("列范围定位失败")
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass
    print()


def test_columns_by_keywords():
    """测试方法2: 关键词定位开始和结束列"""
    print("=== 测试列定位方法2: 关键词定位开始和结束列 ===")
    
    file_path = create_test_excel()
    try:
        operations = [
            {'type': 'columns_by_keywords', 'name': 'cols', 'params': {'header_row': 1, 'start_keyword': '姓名', 'end_keyword': '城市'}}
        ]
        results = locate_batch(file_path, "Sheet1", operations)
        if results.get('cols'):
            print(f"关键词列定位成功: {results['cols']}")
        else:
            print("关键词列定位失败")
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass
    print()


def test_columns_by_start_keyword():
    """测试方法3: 关键词定位开始列，结束列使用最后数据列"""
    print("=== 测试列定位方法3: 关键词定位开始列，结束列使用最后数据列 ===")
    
    file_path = create_test_excel()
    try:
        operations = [
            {'type': 'columns_by_start_keyword', 'name': 'cols', 'params': {'header_row': 1, 'start_keyword': '年龄'}}
        ]
        results = locate_batch(file_path, "Sheet1", operations)
        if results.get('cols'):
            print(f"开始关键词列定位成功: {results['cols']}")
        else:
            print("开始关键词列定位失败")
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass
    print()


def run_all_tests():
    """运行所有测试"""
    print("开始测试列定位功能")
    print("=" * 60)
    
    test_columns_by_range()
    test_columns_by_keywords()
    test_columns_by_start_keyword()
    
    print("所有列定位测试完成！")


if __name__ == "__main__":
    run_all_tests()
