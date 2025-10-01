"""
测试区域定位器功能
"""

import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import pandas as pd
import tempfile
from xlgrab.locator import (
    locate_range, locate_anchor, locate_floating, locate_regions,
    locate_simple, locate_by_text, locate_to_end, get_region_data, get_region_coords
)


def create_test_excel():
    """创建测试Excel文件"""
    # 创建测试数据
    data = {
        'A': ['姓名', 'Alice', 'Bob', 'Charlie', '总计', ''],
        'B': ['年龄', 25, 30, 35, 90, ''],
        'C': ['城市', 'New York', 'London', 'Tokyo', '--', ''],
        'D': ['部门', 'IT', 'HR', 'Finance', '--', ''],
    }
    df = pd.DataFrame(data)
    
    # 创建临时Excel文件
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        df.to_excel(tmp.name, index=False, header=False, engine='openpyxl')
        return tmp.name


def test_locate_range():
    """测试固定区域定位"""
    print("=== 测试 locate_range ===")
    
    file_path = create_test_excel()
    try:
        # 测试固定区域
        region = locate_range(file_path, "Sheet1", "A1:C3")
        if region:
            print(f"定位成功: {region}")
            print(f"坐标: {get_region_coords(region)}")
            
            # 获取数据
            data = get_region_data(file_path, region)
            if data is not None:
                print(f"数据形状: {data.shape}")
                print(f"数据: {data.values.tolist()}")
        else:
            print("定位失败")
        print()
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass


def test_locate_anchor():
    """测试锚点定位"""
    print("=== 测试 locate_anchor ===")
    
    file_path = create_test_excel()
    try:
        # 通过锚点定位
        region = locate_anchor(file_path, "Sheet1", "A", "姓名", 1, (1, 0))
        if region:
            print(f"锚点定位成功: {region}")
            print(f"坐标: {get_region_coords(region)}")
            
            # 获取数据
            data = get_region_data(file_path, region)
            if data is not None:
                print(f"数据形状: {data.shape}")
                print(f"数据: {data.values.tolist()}")
        else:
            print("锚点定位失败")
        print()
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass


def test_locate_floating():
    """测试浮动区域定位"""
    print("=== 测试 locate_floating ===")
    
    file_path = create_test_excel()
    try:
        # 浮动区域定位
        region = locate_floating(file_path, "Sheet1", "A2", "总计", -1)
        if region:
            print(f"浮动定位成功: {region}")
            print(f"坐标: {get_region_coords(region)}")
            
            # 获取数据
            data = get_region_data(file_path, region)
            if data is not None:
                print(f"数据形状: {data.shape}")
                print(f"数据: {data.values.tolist()}")
        else:
            print("浮动定位失败")
        print()
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass


def test_locate_regions():
    """测试批量区域定位"""
    print("=== 测试 locate_regions ===")
    
    file_path = create_test_excel()
    try:
        # 批量定位
        specs = [
            {"sheet": "Sheet1", "area": "A1:C2"},
            {"sheet": "Sheet1", "type": "anchor", "column": "A", "text": "姓名", "occurrence": 1, "offset": (1, 0)},
            {"sheet": "Sheet1", "type": "floating", "start_cell": "A2", "end_keyword": "总计", "end_offset": -1}
        ]
        
        regions = locate_regions(file_path, specs)
        print(f"批量定位成功: {len(regions)} 个区域")
        
        for i, region in enumerate(regions, 1):
            print(f"区域{i}: {region}")
            data = get_region_data(file_path, region)
            if data is not None:
                print(f"  数据形状: {data.shape}")
        print()
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass


def test_convenience_functions():
    """测试便捷函数"""
    print("=== 测试便捷函数 ===")
    
    file_path = create_test_excel()
    try:
        # 简单定位
        region1 = locate_simple(file_path, "Sheet1", "A1:C3")
        print(f"简单定位: {region1}")
        
        # 文本定位
        region2 = locate_by_text(file_path, "Sheet1", "A", "姓名")
        print(f"文本定位: {region2}")
        
        # 定位到最后
        region3 = locate_to_end(file_path, "Sheet1", "A2")
        print(f"定位到最后: {region3}")
        print()
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass


def test_special_areas():
    """测试特殊区域语法"""
    print("=== 测试特殊区域语法 ===")
    
    file_path = create_test_excel()
    try:
        # 测试特殊区域
        areas = ["A1:C5", "A1:last", "A1:lastcol", "A1:lastlast"]
        
        for area in areas:
            region = locate_range(file_path, "Sheet1", area)
            if region:
                print(f"区域 {area}: {region}")
            else:
                print(f"区域 {area}: 定位失败")
        print()
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass


def run_all_tests():
    """运行所有测试"""
    print("开始测试区域定位器")
    print("=" * 50)
    
    test_locate_range()
    test_locate_anchor()
    test_locate_floating()
    test_locate_regions()
    test_convenience_functions()
    test_special_areas()
    
    print("所有测试完成！")


if __name__ == "__main__":
    run_all_tests()