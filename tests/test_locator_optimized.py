"""
测试优化后的区域定位器功能
"""

import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import pandas as pd
import tempfile
from xlgrab.locator import (
    LocatorConfig, LocatorError, LocatedRegion,
    locate_range, locate_anchor, locate_floating, locate_regions,
    locate_simple, locate_by_text, locate_to_end,
    set_default_config, get_default_config, clear_cache, get_cache_info
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


def test_config_functionality():
    """测试配置功能"""
    print("=== 测试配置功能 ===")
    
    # 创建自定义配置
    config = LocatorConfig(
        default_width=5,
        max_columns=10,
        default_height=3,
        search_column="B",
        enable_cache=True
    )
    
    # 设置默认配置
    set_default_config(config)
    retrieved_config = get_default_config()
    
    print(f"默认宽度: {retrieved_config.default_width}")
    print(f"最大列数: {retrieved_config.max_columns}")
    print(f"搜索列: {retrieved_config.search_column}")
    print()


def test_error_handling():
    """测试错误处理"""
    print("=== 测试错误处理 ===")
    
    file_path = create_test_excel()
    try:
        # 测试无效文件路径
        try:
            locate_simple("", "Sheet1", "A1:C3")
        except LocatorError as e:
            print(f"捕获到预期错误: {e.message}")
        
        # 测试无效区域格式
        try:
            locate_simple(file_path, "Sheet1", "A1C3")  # 缺少冒号
        except LocatorError as e:
            print(f"捕获到预期错误: {e.message}")
        
        # 测试无效occurrence
        try:
            locate_anchor(file_path, "Sheet1", "A", "姓名", occurrence=0)
        except LocatorError as e:
            print(f"捕获到预期错误: {e.message}")
        
        # 测试不存在的文本
        try:
            locate_anchor(file_path, "Sheet1", "A", "不存在的文本", occurrence=1)
        except LocatorError as e:
            print(f"捕获到预期错误: {e.message}")
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass
    print()


def test_region_validation():
    """测试区域验证"""
    print("=== 测试区域验证 ===")
    
    file_path = create_test_excel()
    try:
        # 测试有效区域
        region = locate_simple(file_path, "Sheet1", "A1:C3")
        if region:
            print(f"区域有效性: {region.is_valid()}")
            print(f"区域宽度: {region.width()}")
            print(f"区域高度: {region.height()}")
        
        # 测试无效区域（通过配置创建）
        config = LocatorConfig(default_width=0, default_height=0)
        region = locate_simple(file_path, "Sheet1", "A1:C3", config)
        if region:
            print(f"无效区域有效性: {region.is_valid()}")
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass
    print()


def test_cache_functionality():
    """测试缓存功能"""
    print("=== 测试缓存功能 ===")
    
    file_path = create_test_excel()
    try:
        # 清空缓存
        clear_cache()
        print(f"清空后缓存信息: {get_cache_info()}")
        
        # 第一次定位
        region1 = locate_simple(file_path, "Sheet1", "A1:C3")
        print(f"第一次定位后缓存信息: {get_cache_info()}")
        
        # 第二次定位（应该使用缓存）
        region2 = locate_simple(file_path, "Sheet1", "A1:C3")
        print(f"第二次定位后缓存信息: {get_cache_info()}")
        
        # 清空缓存
        clear_cache()
        print(f"清空后缓存信息: {get_cache_info()}")
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass
    print()


def test_enhanced_anchor():
    """测试增强的锚点功能"""
    print("=== 测试增强的锚点功能 ===")
    
    file_path = create_test_excel()
    try:
        # 测试自定义配置的锚点定位
        config = LocatorConfig(default_width=3, default_height=2)
        region = locate_anchor(file_path, "Sheet1", "A", "姓名", 1, (1, 0), config)
        
        if region:
            print(f"锚点区域: {region}")
            print(f"区域宽度: {region.width()}")
            print(f"区域高度: {region.height()}")
            print(f"上下文: {region.context}")
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass
    print()


def test_enhanced_floating():
    """测试增强的浮动区域功能"""
    print("=== 测试增强的浮动区域功能 ===")
    
    file_path = create_test_excel()
    try:
        # 测试自定义搜索列的浮动定位
        config = LocatorConfig(search_column="B", default_width=4)
        region = locate_floating(file_path, "Sheet1", "A2", "90", -1, config)
        
        if region:
            print(f"浮动区域: {region}")
            print(f"区域宽度: {region.width()}")
            print(f"区域高度: {region.height()}")
            print(f"上下文: {region.context}")
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass
    print()


def test_batch_processing():
    """测试批量处理"""
    print("=== 测试批量处理 ===")
    
    file_path = create_test_excel()
    try:
        # 测试批量定位，包含一些会失败的规范
        specs = [
            {"sheet": "Sheet1", "area": "A1:C2"},  # 成功
            {"sheet": "Sheet1", "type": "anchor", "column": "A", "text": "姓名", "occurrence": 1, "offset": (1, 0)},  # 成功
            {"sheet": "Sheet1", "type": "floating", "start_cell": "A2", "end_keyword": "总计", "end_offset": -1},  # 成功
            {"sheet": "Sheet1", "area": "A1C2"},  # 失败：格式错误
            {"sheet": "Sheet1", "type": "anchor", "column": "A", "text": "不存在的文本", "occurrence": 1},  # 失败：文本不存在
            {"sheet": "Sheet1", "type": "unknown"},  # 失败：未知类型
        ]
        
        regions = locate_regions(file_path, specs)
        print(f"批量定位结果: {len(regions)} 个成功区域")
        
        for i, region in enumerate(regions, 1):
            print(f"区域{i}: {region}")
        
    finally:
        try:
            os.unlink(file_path)
        except:
            pass
    print()


def run_all_tests():
    """运行所有测试"""
    print("开始测试优化后的区域定位器")
    print("=" * 60)
    
    test_config_functionality()
    test_error_handling()
    test_region_validation()
    test_cache_functionality()
    test_enhanced_anchor()
    test_enhanced_floating()
    test_batch_processing()
    
    print("所有测试完成！")


if __name__ == "__main__":
    run_all_tests()
