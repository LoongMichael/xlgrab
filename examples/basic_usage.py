"""
xlgrab 基本使用示例
"""

import pandas as pd
import numpy as np
import xlgrab  # 导入后自动注册扩展方法

def main():
    print("=" * 60)
    print("xlgrab 基本使用示例")
    print("=" * 60)
    
    # 创建示例数据
    print("\n1. 创建示例数据:")
    df = pd.DataFrame({
        'name': ['Alice', 'Bob', 'Charlie', 'David', 'Eve', 'Frank', 'Grace', 'Henry'],
        'age': [25, 30, 35, 28, 32, 27, 29, 31],
        'salary': [50000, 60000, 70000, 55000, 65000, 58000, 62000, 68000],
        'department': ['IT', 'HR', 'IT', 'Finance', 'Marketing', 'IT', 'HR', 'Finance']
    })
    
    print("原始数据:")
    print(df)
    
    # 测试扩展方法是否已注册
    print("\n2. 检查扩展方法注册状态:")
    print("DataFrame是否有quick_info方法:", hasattr(df, 'quick_info'))
    print("DataFrame是否有data_profile方法:", hasattr(df, 'data_profile'))
    print("DataFrame是否有filter_by_value方法:", hasattr(df, 'filter_by_value'))
    print("DataFrame是否有find_idx方法:", hasattr(df, 'find_idx'))
    print("Series是否有find_idx方法:", hasattr(df['name'], 'find_idx'))
    
    # 测试直接使用pd.DataFrame（现在已经是增强版本）
    print("\n3. 使用pd.DataFrame（增强版本）:")
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
    
    # 也可以直接使用xlgrab.XlDataFrame
    print("\n4. 使用xlgrab.XlDataFrame:")
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
    
    # 测试find_idx功能
    print("\n7. 测试find_idx功能:")
    try:
        # 测试Series的find_idx
        print("Series的find_idx:")
        print("  查找'Alice'的位置:", df['name'].find_idx('Alice', mode='exact', nth=1))
        print("  查找'Bob'的位置:", df['name'].find_idx('Bob', mode='exact', nth=1))
        print("  查找所有'Alice'的位置:", df['name'].find_idx('Alice', mode='exact', nth=None))
        
        # 测试DataFrame的find_idx
        print("DataFrame的find_idx:")
        print("  按列查找:")
        print("    在'name'列中查找'Alice'的位置:", df.find_idx('name', 'Alice', axis='column', nth=1))
        print("    在'department'列中查找'IT'的所有位置:", df.find_idx('department', 'IT', axis='column', nth=None))
        print("    在'department'列中查找'IT'的第2次位置:", df.find_idx('department', 'IT', axis='column', nth=2))
        print("    在'age'列中查找25的位置:", df.find_idx('age', 25, axis='column', nth=1))
        print("  按行查找:")
        print("    在第0行中查找'Alice'的位置:", df.find_idx(0, 'Alice', axis='row', nth=1))
        print("    在第0行中查找'IT'的位置:", df.find_idx(0, 'IT', axis='row', nth=1))
        print("    在第0行中查找25的位置:", df.find_idx(0, 25, axis='row', nth=1))
        
        # 测试Excel区间转换（使用openpyxl）
        print("Excel区间转换（使用openpyxl）:")
        print("  基本区间转换:")
        print("    df.excel_range('B2:D4'):")
        print(df.excel_range('B2:D4'))
        print("  带header的区间转换:")
        print("    df.excel_range('A1:C3', header=True):")
        print(df.excel_range('A1:C3', header=True))
        print("  带header和index_col的区间转换:")
        print("    df.excel_range('A1:C3', header=True, index_col=0):")
        print(df.excel_range('A1:C3', header=True, index_col=0))
        print("  多区域合并:")
        print("    df.excel_range('A1:C2', 'A3:C4'):")
        print(df.excel_range('A1:C2', 'A3:C4'))
        
        # 测试偏移区间功能
        print("偏移区间功能:")
        print("  统一偏移模式:")
        print("    df.offset_range(1, 3, 2, 4, offset_rows=1, offset_cols=-1):")
        print(df.offset_range(1, 3, 2, 4, offset_rows=1, offset_cols=-1))
        print("  分别偏移模式:")
        print("    df.offset_range(1, 3, 2, 4, offset_start_row=1, offset_end_row=2, offset_start_col=-1, offset_end_col=1):")
        print(df.offset_range(1, 3, 2, 4, offset_start_row=1, offset_end_row=2, offset_start_col=-1, offset_end_col=1))
        print("  自动裁剪模式:")
        print("    df.offset_range(1, 3, 2, 4, offset_rows=10, clip_to_bounds=True):")
        print(df.offset_range(1, 3, 2, 4, offset_rows=10, clip_to_bounds=True))
        
        # 通过 select_range（DSL）完成各种组合的范围选择
        print("通过 select_range（DSL）完成范围选择:")
        print("  示例: 从 A2 到部门为 David 所在行、到第 3 列")
        try:
            print(df.select_range(start='A2', end_row=("find-row", "name", "David"), end_col=("col", 3)))
        except Exception as e:
            print("  选择失败:", e)
        
        # 测试 select_range（DSL）
        print("\n8. 测试 select_range（DSL）:")
        try:
            # 示例1：从 A2 开始，结束列为第6列，结束行匹配 'Finance'
            print("示例1: start='A2', end_col=('col', 6), end_row=('find-row','department','Finance')")
            print(df.select_range(start='A2', end_col=("col", 6), end_row=("find-row", "department", "Finance")))

            # 示例2：从第3行到末行，从 B 列到包含 'IT' 的列（在第0行查找列名场景用法）
            # 这里用 find-col 的 target 指定行为 0（第0行），在该行里寻找匹配文本的列索引
            # 如果你的列名在其他行，可把 target 换成那一行的行索引或标签
            print("\n示例2: start_row=('row',3), end_row=('row','end'), start_col=('col','B'), end_col=('find-col',0,'age')")
            print(df.select_range(start_row=("row", 3), end_row=("row", "end"), start_col=("col", "B"), end_col=("find-col", 0, "age")))

            # 示例3：使用 start/end 同时给端点；end 用 'end' 指定到末行
            print("\n示例3: start='B2', end=('row','end')")
            print(df.select_range(start='B2', end=("row", "end")))

            # 示例4：仅指定列范围（字母列与数字列混用），行默认全行
            print("\n示例4: 仅列范围，start_col=('col','B'), end_col=('col',3)")
            print(df.select_range(start_col=("col", "B"), end_col=("col", 3)))

            # 示例5：仅指定行范围（数字与 'end' 混用），列默认全列
            print("\n示例5: 仅行范围，start_row=('row',2), end_row=('row','end')")
            print(df.select_range(start_row=("row", 2), end_row=("row", "end")))

            # 示例6：用 cell 指定两个端点（等价于 Excel 框选），忽略未给维度用默认
            print("\n示例6: start=('cell','A2'), end=('cell','C4')")
            print(df.select_range(start=("cell", "A2"), end=("cell", "C4")))

            # 示例7：find-row 精细控制模式/第 n 次
            print("\n示例7: end_row=('find-row','department','IT',{'mode':'exact','nth':2})")
            print(df.select_range(start_row=("row", 1), end_row=("find-row", "department", "IT", {"mode": "exact", "nth": 2})))

            # 示例8：find-col 使用 contains/regex 模式
            print("\n示例8: end_col contains → ('find-col',0,'na',{'mode':'contains'})")
            print(df.select_range(start_col=("col", 1), end_col=("find-col", 0, "na", {"mode": "contains"})))
            print("示例8b: end_col regex → ('find-col',0,'^sa',{'mode':'regex'})")
            print(df.select_range(start_col=("col", 1), end_col=("find-col", 0, "^sa", {"mode": "regex"})))

            # 示例9：混合 - 起点用 cell，行终点用 find-row，列终点用字母
            print("\n示例9: start='A2', end_row=('find-row','department','HR'), end_col=('col','C')")
            print(df.select_range(start="A2", end_row=("find-row", "department", "HR"), end_col=("col", "C")))

            # 示例10：clip=False 展示错误（这里故意越界以触发异常）
            print("\n示例10: clip=False 越界演示（期望抛异常）")
            try:
                print(df.select_range(start_row=9999, end_row=10000, clip=False))
            except Exception as e:
                print("  触发异常:", e)

            # 示例11：使用数字行列（1 基）与 'end' 搭配
            print("\n示例11: start_row=1, start_col=1, end_row='end', end_col=('col','D')")
            print(df.select_range(start_row=1, start_col=1, end_row=("row", "end"), end_col=("col", "D")))

            # 示例12：仅给 start，自动补齐到末尾（默认 clip=True）
            print("\n示例12: 仅给 start='C3' → 从C3到表尾")
            print(df.select_range(start="C3"))

            # 示例13：select_range 集成偏移 - 统一偏移
            print("\n示例13: select_range + 统一偏移 offset_rows=1, offset_cols=-1")
            print(df.select_range(start='A2', end=('cell','C4'), offset_rows=1, offset_cols=-1))

            # 示例14：select_range 集成偏移 - 分别偏移
            print("\n示例14: select_range + 分别偏移 offset_start_row=1, offset_end_row=2, offset_start_col=-1, offset_end_col=0")
            print(df.select_range(start='A2', end=('cell','C4'), offset_start_row=1, offset_end_row=2, offset_start_col=-1, offset_end_col=0))
        except Exception as e:
            print(f"select_range 测试失败: {e}")
        except Exception as e:
            print(f"select_range 测试失败: {e}")
        
        # 测试contains模式
        print("contains模式:")
        print("  Series: 查找包含'li'的位置:", df['name'].find_idx('li', mode='contains', nth=1))
        print("  DataFrame: 在'name'列中查找包含'li'的位置:", df.find_idx('name', 'li', mode='contains', nth=1))
        
        # 测试regex模式
        print("regex模式:")
        print("  Series: 查找以'A'开头的位置:", df['name'].find_idx('^A', mode='regex', nth=1))
        print("  DataFrame: 在'name'列中查找以'A'开头的位置:", df.find_idx('name', '^A', mode='regex', nth=1))
        
    except Exception as e:
        print(f"find_idx功能测试失败: {e}")
    
    print("\n" + "=" * 60)
    print("基本架构演示完成！")
    print("第一个功能find_idx已添加并测试通过！")
    print("=" * 60)

if __name__ == "__main__":
    main()