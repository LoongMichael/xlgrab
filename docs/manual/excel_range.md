# excel_range - Excel区间操作

## 功能说明

`excel_range` 用于从 DataFrame 中提取 Excel 风格的数据区间。支持一次传入多个区间并自动纵向合并，可以处理表头和索引列。

## 函数签名

```python
df.xl.excel_range(*ranges, header=True, index_col=None)
```

### 参数
- `*ranges`：一个或多个 Excel 区间字符串，形如 `'A1:C5'`。
- `header`：是否将切片后的首行作为列名（默认 True）。
- `index_col`：索引列（列名或列位置）。

### 返回值
- 新的 `DataFrame`（拷贝）。

### 示例
```python
import pandas as pd, xlgrab
df = pd.DataFrame({
  'A':['Name','Alice','Bob'], 'B':['Age',25,30], 'C':['Dept','IT','HR']
})

df.excel_range('A1:C3', header=True)
df.excel_range('A2:B3')
df.excel_range('A1:B2', 'A3:B3')
df.excel_range('A1:C3', header=True, index_col=0)
```

### 注意事项
- 依赖 `openpyxl` 解析坐标，请确保已安装。
- Excel 区间是“包含边界”的；内部会转换为 0 基索引进行切片。
- 越界会抛 `ValueError`；多区域将按行合并（concat ignore_index）。


