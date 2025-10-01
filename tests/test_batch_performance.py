"""
测试批量定位性能优势
"""

import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import pandas as pd
import tempfile
import time
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


def test_individual_calls():
    """测试单独调用（多次读取）"""
    print("=== 测试单独调用（多次读取） ===")
    
    file_path = create_test_excel()
    try:
        start_time = time.time()
        
        # 多次单独调用，每次都会读取sheet
        operations1 = [{'type': 'rows_by_range', 'name': 'rows1', 'params': {'area': 'A2:D5'}}]
        rows1 = locate_batch(file_path, "Sheet1", operations1).get('rows1')
        
        operations2 = [{'type': 'rows_by_keywords', 'name': 'rows2', 'params': {'start_col': 'A', 'start_keyword': '姓名', 'end_col': 'A', 'end_keyword': '总计'}}]
        rows2 = locate_batch(file_path, "Sheet1", operations2).get('rows2')
        
        operations3 = [{'type': 'columns_by_range', 'name': 'cols1', 'params': {'area': 'A1:D1'}}]
        cols1 = locate_batch(file_path, "Sheet1", operations3).get('cols1')
        
        operations4 = [{'type': 'columns_by_keywords', 'name': 'cols2', 'params': {'header_row': 1, 'start_keyword': '姓名', 'end_keyword': '城市'}}]
        cols2 = locate_batch(file_path, "Sheet1", operations4).get('cols2')
        
        end_time = time.time()
        
        print(f"单独调用结果:")
        print(f"  rows1: {rows1}")
        print(f"  rows2: {rows2}")
        print(f"  cols1: {cols1}")
        print(f"  cols2: {cols2}")
        print(f"耗时: {(end_time - start_time)*1000:.2f}ms")
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass
    print()


def test_batch_calls():
    """测试批量调用（一次读取）"""
    print("=== 测试批量调用（一次读取） ===")
    
    file_path = create_test_excel()
    try:
        start_time = time.time()
        
        # 批量调用，只读取一次sheet
        operations = [
            {
                'type': 'rows_by_range',
                'name': 'rows1',
                'params': {'area': 'A2:D5'}
            },
            {
                'type': 'rows_by_keywords',
                'name': 'rows2',
                'params': {'start_col': 'A', 'start_keyword': '姓名', 'end_col': 'A', 'end_keyword': '总计'}
            },
            {
                'type': 'columns_by_range',
                'name': 'cols1',
                'params': {'area': 'A1:D1'}
            },
            {
                'type': 'columns_by_keywords',
                'name': 'cols2',
                'params': {'header_row': 1, 'start_keyword': '姓名', 'end_keyword': '城市'}
            }
        ]
        
        results = locate_batch(file_path, "Sheet1", operations)
        
        end_time = time.time()
        
        print(f"批量调用结果:")
        for name, result in results.items():
            print(f"  {name}: {result}")
        print(f"耗时: {(end_time - start_time)*1000:.2f}ms")
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass
    print()


def test_performance_comparison():
    """性能对比测试"""
    print("=== 性能对比测试 ===")
    
    file_path = create_test_excel()
    try:
        # 单独调用测试
        start_time = time.time()
        for _ in range(10):  # 重复10次
            locate_batch(file_path, "Sheet1", [{'type': 'rows_by_range', 'name': 'rows1', 'params': {'area': 'A2:D5'}}])
            locate_batch(file_path, "Sheet1", [{'type': 'rows_by_keywords', 'name': 'rows2', 'params': {'start_col': 'A', 'start_keyword': '姓名', 'end_col': 'A', 'end_keyword': '总计'}}])
            locate_batch(file_path, "Sheet1", [{'type': 'columns_by_range', 'name': 'cols1', 'params': {'area': 'A1:D1'}}])
            locate_batch(file_path, "Sheet1", [{'type': 'columns_by_keywords', 'name': 'cols2', 'params': {'header_row': 1, 'start_keyword': '姓名', 'end_keyword': '城市'}}])
        individual_time = time.time() - start_time
        
        # 批量调用测试
        start_time = time.time()
        operations = [
            {'type': 'rows_by_range', 'name': 'rows1', 'params': {'area': 'A2:D5'}},
            {'type': 'rows_by_keywords', 'name': 'rows2', 'params': {'start_col': 'A', 'start_keyword': '姓名', 'end_col': 'A', 'end_keyword': '总计'}},
            {'type': 'columns_by_range', 'name': 'cols1', 'params': {'area': 'A1:D1'}},
            {'type': 'columns_by_keywords', 'name': 'cols2', 'params': {'header_row': 1, 'start_keyword': '姓名', 'end_keyword': '城市'}}
        ]
        for _ in range(10):  # 重复10次
            locate_batch(file_path, "Sheet1", operations)
        batch_time = time.time() - start_time
        
        print(f"单独调用总耗时: {individual_time*1000:.2f}ms")
        print(f"批量调用总耗时: {batch_time*1000:.2f}ms")
        print(f"性能提升: {individual_time/batch_time:.1f}x")
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass
    print()


def run_all_tests():
    """运行所有测试"""
    print("开始测试批量定位性能优势")
    print("=" * 60)
    
    test_individual_calls()
    test_batch_calls()
    test_performance_comparison()
    
    print("所有性能测试完成！")


if __name__ == "__main__":
    run_all_tests()
