"""
xlgrab 高级使用示例
"""

import pandas as pd
import numpy as np
import xlgrab

def main():
    print("=" * 60)
    print("xlgrab 高级使用示例")
    print("=" * 60)
    
    # 创建示例数据
    print("\n1. 创建示例数据:")
    df = pd.DataFrame({
        'product': ['A', 'B', 'C', 'D', 'E'] * 20,
        'sales': np.random.normal(1000, 200, 100),
        'region': ['North', 'South', 'East', 'West'] * 25,
        'date': pd.date_range('2023-01-01', periods=100, freq='D')
    })
    
    print(f"数据形状: {df.shape}")
    print("前5行数据:")
    print(df.head())
    
    # 检查可用的扩展方法
    print("\n2. 检查可用的扩展方法:")
    dataframe_methods = [method for method in dir(df) if not method.startswith('_') and callable(getattr(df, method))]
    print(f"DataFrame可用方法数量: {len(dataframe_methods)}")
    
    # 检查Series扩展方法
    print("\n3. 检查Series扩展方法:")
    series = df['sales']
    series_methods = [method for method in dir(series) if not method.startswith('_') and callable(getattr(series, method))]
    print(f"Series可用方法数量: {len(series_methods)}")
    
    # 测试工具函数
    print("\n4. 检查工具函数:")
    try:
        from xlgrab.utils import create_sample_data
        sample_df = create_sample_data(rows=10)
        print("工具函数可用，创建示例数据成功")
        print(f"示例数据形状: {sample_df.shape}")
    except Exception as e:
        print(f"工具函数不可用: {e}")
    
    print("\n" + "=" * 60)
    print("高级架构演示完成！")
    print("现在可以开始添加具体功能了。")
    print("=" * 60)

if __name__ == "__main__":
    main()