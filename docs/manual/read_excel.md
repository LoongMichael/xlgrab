# read_excel - Excel范围读取

## 功能说明

`read_excel` 函数用于读取Excel文件的指定范围数据。支持读取单个或多个范围，可以纵向合并多个范围的数据。

## 函数签名

```python
xlgrab.read_excel(
    file_path: str, 
    sheet_name: Union[str, int] = 0,
    ranges: Optional[Union[str, List[str]]] = None,
    engine: str = 'openpyxl',
    merge_ranges: bool = False,
    **kwargs
) -> Union[pd.DataFrame, Dict[str, pd.DataFrame]]
```

## 参数说明

- **file_path**: Excel文件路径
- **sheet_name**: 工作表名称或索引，默认为0（第一个工作表）
- **ranges**: 单个范围字符串或范围列表，如 "A1:C10" 或 ["A1:C10", "E1:G10"]
- **engine**: 读取引擎，默认 'openpyxl'，也可用 'calamine' 等
- **merge_ranges**: 是否纵向合并多个范围，默认False返回字典
- **kwargs**: 传递给 pd.read_excel 的其他参数

## 返回值

- **单个范围**：返回 DataFrame
- **多个范围且 merge_ranges=False**：返回字典 `{range: DataFrame}`
- **多个范围且 merge_ranges=True**：返回纵向合并后的 DataFrame

## 使用示例

### 1. 读取单个范围

```python
import xlgrab

# 读取单个范围
df = xlgrab.read_excel("data.xlsx", ranges="A1:C10")
print(df.shape)  # 输出: (10, 3)
```

### 2. 读取多个范围（返回字典）

```python
# 读取多个范围，返回字典
dfs = xlgrab.read_excel("data.xlsx", ranges=["A1:C10", "E1:G10"])
print(dfs.keys())  # 输出: dict_keys(['A1:C10', 'E1:G10'])

# 访问特定范围的数据
df1 = dfs['A1:C10']
df2 = dfs['E1:G10']
```

### 3. 读取多个范围并合并

```python
# 读取多个范围并纵向合并
df = xlgrab.read_excel(
    "data.xlsx", 
    ranges=["A1:C10", "A15:C25"], 
    merge_ranges=True
)
print(df.shape)  # 输出: (21, 3) - 10行 + 11行
```

### 4. 指定工作表

```python
# 按名称指定工作表
df = xlgrab.read_excel(
    "data.xlsx", 
    sheet_name="Sheet1", 
    ranges="A1:C10"
)

# 按索引指定工作表
df = xlgrab.read_excel(
    "data.xlsx", 
    sheet_name=1,  # 第二个工作表
    ranges="A1:C10"
)
```

### 5. 使用不同引擎

```python
# 使用 openpyxl 引擎（默认）
df = xlgrab.read_excel("data.xlsx", ranges="A1:C10", engine='openpyxl')

# 使用 calamine 引擎（更快，但功能较少）
df = xlgrab.read_excel("data.xlsx", ranges="A1:C10", engine='calamine')
```

### 6. 结合其他参数

```python
# 传递额外的 pandas 参数
df = xlgrab.read_excel(
    "data.xlsx", 
    ranges="A1:C10",
    engine='openpyxl',
    dtype={'A': 'str', 'B': 'int64'}  # 指定数据类型
)
```

### 7. 结合 xlgrab 的其他方法

```python
# 读取数据并使用 xlgrab 方法处理
df = xlgrab.read_excel("data.xlsx", ranges="A1:C10")

# 使用 apply_header 处理表头
df_processed = df.xl.apply_header(0)

# 使用 find_idx 查找数据
position = df_processed.xl.find_idx('column_name', 'target_value')
```

## 范围格式说明

支持标准的Excel范围格式：

- **单单元格**: "A1", "B5", "Z100"
- **矩形范围**: "A1:C10", "B2:D6"
- **整行**: "1:1", "5:5"
- **整列**: "A:A", "B:D"
- **多范围**: ["A1:C10", "E1:G10", "I1:K10"]

## 处理逻辑

1. **解析范围**：使用 `openpyxl.utils.range_boundaries` 解析范围字符串
2. **计算参数**：根据范围计算 `usecols`、`skiprows`、`nrows` 参数
3. **读取数据**：调用 `pd.read_excel` 读取指定范围
4. **合并处理**：如果指定多个范围且 `merge_ranges=True`，纵向合并数据

## 性能优化

### 1. 选择合适的引擎

```python
# 对于大文件，calamine 引擎通常更快
df = xlgrab.read_excel("large_file.xlsx", ranges="A1:C1000", engine='calamine')
```

### 2. 只读取需要的范围

```python
# 只读取需要的数据范围，而不是整个工作表
df = xlgrab.read_excel("data.xlsx", ranges="A1:C100")  # 只读100行3列
```

### 3. 批量处理多个范围

```python
# 一次性读取多个相关范围，避免多次文件访问
ranges = ["A1:C10", "A15:C25", "A30:C40"]
dfs = xlgrab.read_excel("data.xlsx", ranges=ranges)
```

## 错误处理

```python
try:
    df = xlgrab.read_excel("data.xlsx", ranges="A1:C10")
except ValueError as e:
    print(f"范围格式错误: {e}")
except FileNotFoundError:
    print("文件不存在")
except Exception as e:
    print(f"读取失败: {e}")
```

## 注意事项

- **依赖要求**：需要安装 `openpyxl >= 3.0.0` 或 `calamine`
- **范围有效性**：确保指定的范围在Excel文件中存在
- **数据类型**：读取的数据类型由pandas自动推断，可通过 `dtype` 参数指定
- **内存使用**：大范围数据会占用较多内存，建议分批处理

## 相关函数

- `unmerge_excel()`: 处理Excel合并单元格
- `excel_range()`: DataFrame的Excel区间操作
- `apply_header()`: 表头处理

## 实际应用场景

### 1. 数据提取

```python
# 从复杂Excel文件中提取特定数据区域
df = xlgrab.read_excel(
    "report.xlsx", 
    sheet_name="Summary", 
    ranges="B5:F20"  # 提取汇总数据
)
```

### 2. 多区域数据合并

```python
# 从不同区域读取数据并合并
ranges = ["A1:C10", "A15:C25", "A30:C40"]  # 三个数据区域
df = xlgrab.read_excel("data.xlsx", ranges=ranges, merge_ranges=True)
```

### 3. 批量数据处理

```python
# 批量处理多个文件
files = ["file1.xlsx", "file2.xlsx", "file3.xlsx"]
all_data = []

for file in files:
    df = xlgrab.read_excel(file, ranges="A1:C100")
    all_data.append(df)

# 合并所有数据
combined_df = pd.concat(all_data, ignore_index=True)
```
