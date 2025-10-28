"""
Excel 写入模块

提供向现有Excel文件写入数据的功能
"""

import pandas as pd
import openpyxl
from typing import Union, Optional
import warnings


def write_to_excel(df: pd.DataFrame,
                   excel_name: str,
                   sheet_name: Union[str, int] = 0,
                   start_row: int = 1,
                   start_col: int = 1,
                   end_row: Optional[int] = None,
                   end_col: Optional[int] = None,
                   overwrite: bool = False,
                   header: bool = True,
                   index: bool = False) -> None:
    """
    向现有Excel文件的指定位置写入DataFrame数据
    
    参数:
    df: 要写入的DataFrame
    excel_name: Excel文件路径
    sheet_name: 工作表名称或索引，默认为0（第一个工作表）
    start_row: 起始行位置（从1开始），默认1
    start_col: 起始列位置（从1开始），默认1
    end_row: 结束行位置（从1开始），如果为None则自动计算
    end_col: 结束列位置（从1开始），如果为None则自动计算
    overwrite: 是否覆盖现有数据，默认False（追加模式）
    header: 是否写入列名，默认True
    index: 是否写入行索引，默认False
    
    示例:
        >>> import pandas as pd
        >>> import xlgrab
        >>> 
        >>> # 创建测试数据
        >>> df = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
        >>> 
        >>> # 写入到Excel文件的B2位置
        >>> xlgrab.write_to_excel(df, "test.xlsx", start_row=2, start_col=2)
        >>> 
        >>> # 写入到指定工作表
        >>> xlgrab.write_to_excel(df, "test.xlsx", sheet_name="Sheet2", start_row=1, start_col=1)
    """
    
    # 验证输入参数
    if not isinstance(df, pd.DataFrame):
        raise ValueError("df参数必须是pandas DataFrame")
    
    if not isinstance(excel_name, str):
        raise ValueError("excel_name参数必须是字符串")
    
    if start_row < 1:
        raise ValueError("start_row必须大于等于1")
    
    if start_col < 1:
        raise ValueError("start_col必须大于等于1")
    
    # 计算DataFrame的实际尺寸
    df_rows = len(df)
    df_cols = len(df.columns)
    
    # 计算结束位置
    if end_row is None:
        end_row = start_row + df_rows - 1
    if end_col is None:
        end_col = start_col + df_cols - 1
    
    # 验证结束位置
    if end_row < start_row:
        raise ValueError("end_row不能小于start_row")
    if end_col < start_col:
        raise ValueError("end_col不能小于start_col")
    
    # 检查是否需要调整DataFrame大小
    required_rows = end_row - start_row + 1
    required_cols = end_col - start_col + 1
    
    if not overwrite and (df_rows > required_rows or df_cols > required_cols):
        warnings.warn(
            f"DataFrame尺寸({df_rows}x{df_cols})大于目标区域({required_rows}x{required_cols})，"
            f"数据将被截断。使用overwrite=True来覆盖现有数据。",
            UserWarning
        )
    
    # 调整DataFrame大小以匹配目标区域
    if df_rows != required_rows or df_cols != required_cols:
        # 创建新的DataFrame来填充目标区域
        target_df = pd.DataFrame(index=range(required_rows), columns=range(required_cols))
        
        # 复制数据到目标DataFrame
        for i in range(min(df_rows, required_rows)):
            for j in range(min(df_cols, required_cols)):
                target_df.iloc[i, j] = df.iloc[i, j]
        
        df = target_df
    
    try:
        # 尝试加载现有工作簿
        try:
            workbook = openpyxl.load_workbook(excel_name)
        except FileNotFoundError:
            # 如果文件不存在，创建新的工作簿
            workbook = openpyxl.Workbook()
            # 删除默认工作表
            if 'Sheet' in workbook.sheetnames:
                workbook.remove(workbook['Sheet'])
        
        # 获取或创建工作表
        if isinstance(sheet_name, int):
            if sheet_name < len(workbook.sheetnames):
                worksheet = workbook.worksheets[sheet_name]
            else:
                # 创建新工作表
                worksheet = workbook.create_sheet(f"Sheet{sheet_name + 1}")
        else:
            if sheet_name in workbook.sheetnames:
                worksheet = workbook[sheet_name]
            else:
                # 创建新工作表
                worksheet = workbook.create_sheet(sheet_name)
        
        # 准备写入的数据
        data_to_write = df.values.tolist()
        
        # 如果需要写入列名
        if header and not df.empty:
            # 获取列名
            if isinstance(df.columns, pd.RangeIndex):
                col_names = [f"Column_{i}" for i in df.columns]
            else:
                col_names = df.columns.tolist()
            
            # 在数据上方插入列名
            data_to_write = [col_names] + data_to_write
        
        # 如果需要写入行索引
        if index and not df.empty:
            # 获取行索引
            if isinstance(df.index, pd.RangeIndex):
                row_names = [f"Row_{i}" for i in df.index]
            else:
                row_names = df.index.tolist()
            
            # 在每行数据前插入行索引
            for i, row_data in enumerate(data_to_write):
                data_to_write[i] = [row_names[i]] + row_data
        
        # 写入数据到工作表
        for i, row_data in enumerate(data_to_write):
            for j, cell_value in enumerate(row_data):
                cell_row = start_row + i
                cell_col = start_col + j
                
                # 设置单元格值
                worksheet.cell(row=cell_row, column=cell_col, value=cell_value)
        
        # 保存工作簿
        workbook.save(excel_name)
        
        print(f"成功写入数据到 {excel_name} 的工作表 '{worksheet.title}' "
              f"位置 ({start_row},{start_col}) 到 ({end_row},{end_col})")
        
    except Exception as e:
        raise RuntimeError(f"写入Excel文件失败: {e}")


def write_range_to_excel(data: Union[pd.DataFrame, list, tuple],
                        excel_name: str,
                        sheet_name: Union[str, int] = 0,
                        start_row: int = 1,
                        start_col: int = 1,
                        end_row: Optional[int] = None,
                        end_col: Optional[int] = None) -> None:
    """
    向Excel文件的指定范围写入数据（简化版本）
    
    参数:
    data: 要写入的数据，可以是DataFrame、列表或元组
    excel_name: Excel文件路径
    sheet_name: 工作表名称或索引
    start_row: 起始行位置（从1开始）
    start_col: 起始列位置（从1开始）
    end_row: 结束行位置（从1开始）
    end_col: 结束列位置（从1开始）
    
    示例:
        >>> import xlgrab
        >>> 
        >>> # 写入列表数据
        >>> data = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
        >>> xlgrab.write_range_to_excel(data, "test.xlsx", start_row=2, start_col=2)
        >>> 
        >>> # 写入DataFrame
        >>> df = pd.DataFrame({'A': [1, 2], 'B': [3, 4]})
        >>> xlgrab.write_range_to_excel(df, "test.xlsx", sheet_name="Sheet2")
    """
    
    # 转换数据为DataFrame
    if isinstance(data, pd.DataFrame):
        df = data
    elif isinstance(data, (list, tuple)):
        df = pd.DataFrame(data)
    else:
        raise ValueError("data参数必须是DataFrame、列表或元组")
    
    # 调用主函数
    write_to_excel(
        df=df,
        excel_name=excel_name,
        sheet_name=sheet_name,
        start_row=start_row,
        start_col=start_col,
        end_row=end_row,
        end_col=end_col,
        overwrite=True,
        header=False,
        index=False
    )
