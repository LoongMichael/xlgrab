# unmerge_excel - 合并单元格处理

## 功能说明

`unmerge_excel` 函数用于解开Excel文件中的所有合并单元格并填充值。该函数支持批量处理多个文件和工作表，可以保持或忽略单元格格式。

## 函数签名

```python
xlgrab.unmerge_excel(
    file_path: Union[str, List[str]], 
    output_path: Optional[Union[str, List[str]]] = None,
    sheet_names: Optional[Union[str, List[str]]] = None,
    copy_style: bool = True,
    verbose: bool = False
) -> Dict
```

## 参数说明

- **file_path**: 输入Excel文件路径或文件路径列表
- **output_path**: 输出Excel文件路径或路径列表，如果为None则覆盖原文件
- **sheet_names**: 要处理的工作表名称或名称列表，None表示处理所有工作表
- **copy_style**: 是否复制单元格格式（数字格式、数据类型等），默认True
- **verbose**: 是否显示详细处理信息，默认False

## 返回值

返回一个字典，包含处理结果统计：
- `total_files`: 处理的文件数量
- `total_sheets`: 处理的工作表数量
- `total_merged`: 处理的合并单元格数量
- `files_info`: 各文件的详细信息

## 使用示例

### 1. 处理单个文件

```python
import xlgrab

# 处理单个文件的所有工作表，保存到新文件
result = xlgrab.unmerge_excel("input.xlsx", "output.xlsx")
print(f"处理了 {result['total_merged']} 个合并单元格")

# 直接覆盖原文件
result = xlgrab.unmerge_excel("input.xlsx")
```

### 2. 处理指定工作表

```python
# 只处理指定工作表
result = xlgrab.unmerge_excel(
    "input.xlsx", 
    "output.xlsx", 
    sheet_names="Sheet1"
)

# 处理多个指定工作表
result = xlgrab.unmerge_excel(
    "input.xlsx", 
    "output.xlsx", 
    sheet_names=["Sheet1", "Sheet2"]
)
```

### 3. 批量处理多个文件

```python
# 批量处理多个文件
files = ["file1.xlsx", "file2.xlsx", "file3.xlsx"]
result = xlgrab.unmerge_excel(files)

# 批量处理并指定输出路径
input_files = ["file1.xlsx", "file2.xlsx", "file3.xlsx"]
output_files = ["out1.xlsx", "out2.xlsx", "out3.xlsx"]
result = xlgrab.unmerge_excel(input_files, output_files)
```

### 4. 显示详细处理信息

```python
# 显示详细的处理过程
result = xlgrab.unmerge_excel(
    "input.xlsx", 
    "output.xlsx", 
    verbose=True
)
```

输出示例：
```
============================================================
处理文件: input.xlsx
============================================================
  工作表 'Sheet1': 处理了 5 个合并单元格
  已保存到: output.xlsx
  共处理 5 个合并单元格

============================================================
全部完成! 共处理 1 个文件, 1 个工作表, 5 个合并单元格
============================================================
```

### 5. 只复制值，不复制格式

```python
# 不复制格式，只复制值
result = xlgrab.unmerge_excel(
    "input.xlsx", 
    "output.xlsx", 
    copy_style=False
)
```

## 处理逻辑

1. **收集合并信息**：遍历所有合并单元格，记录其范围和值
2. **取消合并**：调用 `worksheet.unmerge_cells()` 取消所有合并
3. **填充值**：将原合并单元格的值填充到所有相关单元格中
4. **复制格式**（可选）：如果 `copy_style=True`，同时复制数字格式和数据类型

## 注意事项

- **依赖要求**：需要安装 `openpyxl >= 3.0.0`
- **文件安全**：建议先备份重要文件，或使用不同的输出路径
- **磁盘空间**：确保有足够的磁盘空间保存输出文件
- **格式保持**：默认会保持数字格式和数据类型，可通过 `copy_style=False` 关闭
- **错误处理**：如果某个文件处理失败，会记录错误信息并继续处理其他文件

## 错误处理

函数会优雅地处理各种错误情况：

```python
result = xlgrab.unmerge_excel("input.xlsx", "output.xlsx", verbose=True)

# 如果文件不存在或格式错误，会在 files_info 中记录错误
for file_info in result['files_info']:
    if 'error' in file_info:
        print(f"文件 {file_info['input_file']} 处理失败: {file_info['error']}")
```

## 性能考虑

- 对于大文件，处理时间主要取决于合并单元格的数量
- 使用 `verbose=False`（默认）可以获得更好的性能
- 批量处理时，函数会逐个处理文件，避免内存占用过大

## 相关函数

- `unmerge_sheet()`: 处理单个工作表的合并单元格
- `read_excel_range()`: 读取Excel文件的指定范围数据
