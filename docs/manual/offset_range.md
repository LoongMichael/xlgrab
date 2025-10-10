# offset_range - 偏移范围选择

## 功能说明

`offset_range` 用于以 1 基行列坐标选择一个区域，并在选择后应用统一或分别的偏移。支持自动裁剪到边界，提供灵活的范围选择功能。

## 函数签名

```python
df.xl.offset_range(start_row, end_row, start_col, end_col,
                   offset_rows=0, offset_cols=0,
                   offset_start_row=None, offset_end_row=None,
                   offset_start_col=None, offset_end_col=None,
                   clip_to_bounds=False)
```

### 参数要点
- `start_row/end_row/start_col/end_col`：均为 1 基（A=1）。
- 统一偏移：`offset_rows/offset_cols`；分别偏移：四个 `offset_*` 参数（二者互斥，后者优先）。
- `clip_to_bounds`：True 则越界自动裁剪；False 则越界抛错。

### 返回值
- 偏移并裁剪后的 `DataFrame`（拷贝）。

### 示例
```python
import pandas as pd, xlgrab
df = pd.DataFrame({'A':['A1','A2','A3','A4'], 'B':['B1','B2','B3','B4'], 'C':['C1','C2','C3','C4']})

# 统一偏移：行+1，列-1
df.offset_range(1, 3, 2, 3, offset_rows=1, offset_cols=-1)

# 分别偏移
df.offset_range(1, 3, 2, 3, offset_start_row=1, offset_end_row=2, offset_start_col=-1, offset_end_col=0)

# 自动裁剪
df.offset_range(1, 3, 2, 3, offset_rows=10, clip_to_bounds=True)
```

### 常见问题
- 当 `clip_to_bounds=False` 且结果越界时会抛错；为避免抛错可开启裁剪。


