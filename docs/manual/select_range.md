## select_range 使用手册（区间 DSL）

统一表达单元格、行/列、查找端点与偏移，内部最终复用 `offset_range` 执行切片。

### 签名

```python
df.select_range(*, start=None, end=None, start_row=None, end_row=None, start_col=None, end_col=None,
                clip=True,
                offset_rows=0, offset_cols=0,
                offset_start_row=None, offset_end_row=None,
                offset_start_col=None, offset_end_col=None)
```

### 端点规范
- 字符串：`'A2'`（cell）、`'F'`（列）。
- 整数：1 基行/列索引。
- 元组：
  - `('cell','A2')`
  - `('row', 10)`
  - `('col', 'F'|6)`
  - `('find-row', target, q, {mode, nth, na, flags})`
  - `('find-col', targetRowIndexOrLabel, q, {mode, nth, na, flags})`

### 默认与优先级
- 未给出的边界默认：`start_row=1, start_col=1, end_row=末行, end_col=末列`。
- `start_row/col/end_row/col` 会覆盖前述推断。
- `clip=True` 时会自动裁剪到范围内；否则越界抛错。

### 示例
```python
import pandas as pd, xlgrab
df = pd.DataFrame({
  'name':['Alice','Bob','Charlie','David'],
  'age':[25,30,35,28],
  'dept':['IT','HR','IT','Finance']
})

# 1) 单元格到单元格
df.select_range(start='A2', end=('cell','C3'))

# 2) 指定行范围到表尾
df.select_range(start_row=('row', 2))

# 3) 指定列范围（字母+数字混合）
df.select_range(start_col='B', end_col=('col', 3), start_row=2, end_row=3)

# 4) 使用 find-row/col
df.select_range(
  start_row=('find-row','name','Alice', {'mode':'exact','nth':1}),
  end_row=('find-row','name','David', {'mode':'exact','nth':1}),
  start_col=('find-col', 0, 'age', {'mode':'exact'}),
  end_col=('find-col', 0, 'dept', {'mode':'exact'})
)

# 5) 仅指定起点，自动到底（省略末端）
df.select_range(start='B2')

# 6) 集成偏移
df.select_range(start='A2', end=('cell','C3'), offset_rows=1, offset_cols=-1, clip=True)
```

### 提示
- `'find-col'` 中的 `target` 是行索引或标签，常用 `0` 表示第一行。


