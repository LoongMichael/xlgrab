"""
xlgrab 演示脚本
"""

import pandas as pd
import numpy as np
import xlgrab  # 导入后自动注册扩展方法

def main():
    print("=" * 60)
    print("xlgrab 演示")
    print("=" * 60)
    
    # 创建示例数据
    print("\n1. 创建示例数据:")
    df = pd.DataFrame({
        'name': ['Alice', 'Bob', 'Charlie', 'David', 'Eve'],
        'age': [25, 30, 35, 28, 32],
        'salary': [50000, 60000, 70000, 55000, 65000],
        'department': ['IT', 'HR', 'IT', 'Finance', 'Marketing']
    })
    
    print("原始数据:")
    print(df)
    
    # 测试扩展方法注册状态
    print("\n2. 检查扩展方法注册状态:")
    print("DataFrame是否有quick_info方法:", hasattr(df, 'quick_info'))
    print("DataFrame是否有data_profile方法:", hasattr(df, 'data_profile'))
    print("DataFrame是否有filter_by_value方法:", hasattr(df, 'filter_by_value'))
    
    # 测试直接使用pd.DataFrame（现在已经是增强版本）
    print("\n3. 测试pd.DataFrame（增强版本）:")
    try:
        # 现在pd.DataFrame就是我们的增强版本
        enhanced_df = pd.DataFrame({
            'name': ['Alice', 'Bob', 'Charlie'],
            'age': [25, 30, 35],
            'salary': [50000, 60000, 70000]
        })
        print("增强版DataFrame创建成功:")
        print(enhanced_df)
        print(f"数据类型: {type(enhanced_df)}")
    except Exception as e:
        print(f"增强版DataFrame创建失败: {e}")
    
    # 测试Series（现在也是增强版本）
    print("\n4. 测试Series（增强版本）:")
    try:
        series = enhanced_df['age']
        print(f"Series类型: {type(series)}")
        print(f"Series值: {series.tolist()}")
    except Exception as e:
        print(f"Series测试失败: {e}")
    
    # 也可以直接使用xlgrab.XlDataFrame
    print("\n5. 测试xlgrab.XlDataFrame:")
    try:
        xl_df = xlgrab.XlDataFrame({
            'name': ['Alice', 'Bob', 'Charlie'],
            'age': [25, 30, 35],
            'salary': [50000, 60000, 70000]
        })
        print("XlDataFrame创建成功:")
        print(xl_df)
        print(f"数据类型: {type(xl_df)}")
    except Exception as e:
        print(f"XlDataFrame创建失败: {e}")
    
    # 检查工具函数
    print("\n6. 检查工具函数:")
    try:
        from xlgrab.utils import create_sample_data
        sample_df = create_sample_data(rows=5)
        print("工具函数可用")
    except Exception as e:
        print(f"工具函数不可用: {e}")
    
    print("\n" + "=" * 60)
    print("架构演示完成！")
    print("现在可以开始一个一个添加功能了。")
    print("=" * 60)

if __name__ == "__main__":
    main()