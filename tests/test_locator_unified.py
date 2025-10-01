"""
测试统一架构的区域定位功能
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
        'A': ['标题', '姓名', 'Alice', 'Bob', 'Charlie', '总计', ''],
        'B': ['类型', '年龄', 25, 30, 35, 90, ''],
        'C': ['位置', '城市', 'New York', 'London', 'Tokyo', '--', ''],
        'D': ['部门', '部门', 'IT', 'HR', 'Finance', '--', ''],
    }
    df = pd.DataFrame(data)
    
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        df.to_excel(tmp.name, index=False, header=False, engine='openpyxl')
        return tmp.name


def test_method1_direct_range():
    """测试方法1: 直接指定区域"""
    print("=== 测试方法1: 直接指定区域 ===")
    
    file_path = create_test_excel()
    try:
        operations = [
            {'type': 'rows_by_range', 'name': 'rows', 'params': {'area': 'A2:D5'}}
        ]
        results = locate_batch(file_path, "Sheet1", operations)
        if results.get('rows'):
            print(f"直接区域定位成功: {results['rows']}")
        else:
            print("直接区域定位失败")
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass
    print()


def test_method2_keywords():
    """测试方法2: 关键词定位开始和结束行"""
    print("=== 测试方法2: 关键词定位开始和结束行 ===")
    
    file_path = create_test_excel()
    try:
        operations = [
            {'type': 'rows_by_keywords', 'name': 'rows1', 'params': {'start_col': 'A', 'start_keyword': '姓名', 'end_col': 'A', 'end_keyword': '总计'}},
            {'type': 'rows_by_keywords', 'name': 'rows2', 'params': {'start_col': 'A', 'start_keyword': '姓名', 'end_col': 'B', 'end_keyword': '90'}}
        ]
        results = locate_batch(file_path, "Sheet1", operations)
        if results.get('rows1'):
            print(f"关键词定位成功: {results['rows1']}")
        else:
            print("关键词定位失败")
        
        if results.get('rows2'):
            print(f"不同列关键词定位成功: {results['rows2']}")
        else:
            print("不同列关键词定位失败")
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass
    print()


def test_method3_start_keyword():
    """测试方法3: 关键词定位开始行，结束行使用最后数据行"""
    print("=== 测试方法3: 关键词定位开始行，结束行使用最后数据行 ===")
    
    file_path = create_test_excel()
    try:
        operations = [
            {'type': 'rows_by_start_keyword', 'name': 'rows', 'params': {'start_col': 'A', 'start_keyword': '姓名'}}
        ]
        results = locate_batch(file_path, "Sheet1", operations)
        if results.get('rows'):
            print(f"开始关键词定位成功: {results['rows']}")
        else:
            print("开始关键词定位失败")
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass
    print()


def test_error_cases():
    """测试错误情况"""
    print("=== 测试错误情况 ===")
    
    file_path = create_test_excel()
    try:
        operations = [
            {'type': 'rows_by_start_keyword', 'name': 'error1', 'params': {'start_col': 'A', 'start_keyword': '不存在的关键词'}},
            {'type': 'rows_by_range', 'name': 'error2', 'params': {'area': 'A2B5'}},
            {'type': 'rows_by_keywords', 'name': 'error3', 'params': {'start_col': 'A', 'start_keyword': '总计', 'end_col': 'A', 'end_keyword': '姓名'}}
        ]
        results = locate_batch(file_path, "Sheet1", operations)
        
        for name, result in results.items():
            if result:
                print(f"{name} 意外成功: {result}")
            else:
                print(f"{name} 正确返回None")
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass
    print()


def run_all_tests():
    """运行所有测试"""
    print("开始测试批量定位功能")
    print("=" * 60)
    
    test_method1_direct_range()
    test_method2_keywords()
    test_method3_start_keyword()
    test_error_cases()
    
    print("所有测试完成！")


if __name__ == "__main__":
    run_all_tests()
