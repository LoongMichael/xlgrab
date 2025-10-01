"""
测试核心区域定位功能
"""

import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import pandas as pd
import tempfile
from xlgrab.locator import (
    locate_range, locate_anchor, locate_floating, LocatedRegion
)


def create_test_excel():
    """创建测试Excel文件"""
    data = {
        'A': ['姓名', 'Alice', 'Bob', 'Charlie', '总计', ''],
        'B': ['年龄', 25, 30, 35, 90, ''],
        'C': ['城市', 'New York', 'London', 'Tokyo', '--', ''],
        'D': ['部门', 'IT', 'HR', 'Finance', '--', ''],
    }
    df = pd.DataFrame(data)
    
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        df.to_excel(tmp.name, index=False, header=False, engine='openpyxl')
        return tmp.name


def test_locate_range():
    """测试固定区域定位"""
    print("=== 测试 locate_range ===")
    
    file_path = create_test_excel()
    try:
        # 测试基本区域
        region = locate_range(file_path, "Sheet1", "A1:C3")
        if region:
            print(f"定位成功: {region}")
            print(f"坐标: sheet={region.sheet}, {region.start_row}:{region.end_row}, {region.start_col}:{region.end_col}")
        else:
            print("定位失败")
        
        # 测试无效区域
        region = locate_range(file_path, "Sheet1", "A1C3")  # 缺少冒号
        if region:
            print(f"意外成功: {region}")
        else:
            print("正确返回None")
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass
    print()


def test_locate_anchor():
    """测试锚点行号定位"""
    print("=== 测试 locate_anchor ===")
    
    file_path = create_test_excel()
    try:
        # 测试精确匹配
        print("--- 精确匹配测试 ---")
        row = locate_anchor(file_path, "Sheet1", "A", "姓名", 1)
        if row:
            print(f"精确匹配'姓名'成功: 行号 {row}")
        else:
            print("精确匹配'姓名'失败")
        
        row = locate_anchor(file_path, "Sheet1", "A", "Alice", 1)
        if row:
            print(f"精确匹配'Alice'成功: 行号 {row}")
        else:
            print("精确匹配'Alice'失败")
        
        # 测试包含匹配
        print("--- 包含匹配测试 ---")
        row = locate_anchor(file_path, "Sheet1", "A", "li", 1, contains=True)
        if row:
            print(f"包含匹配'li'成功: 行号 {row}")
        else:
            print("包含匹配'li'失败")
        
        row = locate_anchor(file_path, "Sheet1", "A", "ob", 1, contains=True)
        if row:
            print(f"包含匹配'ob'成功: 行号 {row}")
        else:
            print("包含匹配'ob'失败")
        
        row = locate_anchor(file_path, "Sheet1", "A", "总计", 1, contains=True)
        if row:
            print(f"包含匹配'总计'成功: 行号 {row}")
        else:
            print("包含匹配'总计'失败")
        
        # 测试不存在的文本
        print("--- 不存在文本测试 ---")
        row = locate_anchor(file_path, "Sheet1", "A", "不存在的文本", 1)
        if row:
            print(f"意外成功: 行号 {row}")
        else:
            print("正确返回None")
        
        row = locate_anchor(file_path, "Sheet1", "A", "xyz", 1, contains=True)
        if row:
            print(f"意外成功: 行号 {row}")
        else:
            print("正确返回None")
        
        # 测试不同列
        print("--- 不同列测试 ---")
        row = locate_anchor(file_path, "Sheet1", "B", "25", 1)
        if row:
            print(f"B列精确匹配'25'成功: 行号 {row}")
        else:
            print("B列精确匹配'25'失败")
        
        row = locate_anchor(file_path, "Sheet1", "B", "5", 1, contains=True)
        if row:
            print(f"B列包含匹配'5'成功: 行号 {row}")
        else:
            print("B列包含匹配'5'失败")
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass
    print()


def test_locate_floating():
    """测试浮动区域定位"""
    print("=== 测试 locate_floating ===")
    
    file_path = create_test_excel()
    try:
        # 测试基本浮动定位
        region = locate_floating(file_path, "Sheet1", "A2", "总计", -1)
        if region:
            print(f"浮动定位成功: {region}")
            print(f"坐标: sheet={region.sheet}, {region.start_row}:{region.end_row}, {region.start_col}:{region.end_col}")
        else:
            print("浮动定位失败")
        
        # 测试到最后数据行
        region = locate_floating(file_path, "Sheet1", "A2")
        if region:
            print(f"到最后数据行成功: {region}")
        else:
            print("到最后数据行失败")
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass
    print()


def test_located_region():
    """测试LocatedRegion类"""
    print("=== 测试 LocatedRegion ===")
    
    region = LocatedRegion("Sheet1", 1, 3, 1, 3)
    print(f"创建区域: {region}")
    print(f"字符串表示: {str(region)}")
    print(f"repr表示: {repr(region)}")
    print()


def run_all_tests():
    """运行所有测试"""
    print("开始测试核心区域定位功能")
    print("=" * 50)
    
    test_located_region()
    test_locate_range()
    test_locate_anchor()
    test_locate_floating()
    
    print("所有测试完成！")


if __name__ == "__main__":
    run_all_tests()
