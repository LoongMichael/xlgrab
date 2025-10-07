## find_idx 使用手册

查找位置（按列或按行），支持三种模式与第 n 次命中选择。

### 签名（DataFrame）

```python
df.find_idx(target, q, mode='exact', na=False, nth=1, axis='column')
```

### 签名（Series）

```python
s.find_idx(q, mode='exact', na=False, nth=1)
```

### 参数
- `target`：DataFrame 查找目标。按列时是列名或列索引；按行时是行索引（整数/标签）。
- `q`：查询值或正则表达式。
- `mode`：`'exact'|'contains'|'regex'`。
- `na`：`contains/regex` 下传入 `.str.contains` 的 `na`。
- `flags`：`regex` 模式下的 `re` 标志位（在 DataFrame 版本可通过 `select_range` 的 find-* 使用）。
- `nth`：`None` 返回全部命中索引（ndarray）；正数第 n 个；负数从末尾数（-1 为最后一次）。
- `axis`：`'column'|'row'`，默认按列。

### 返回值
- `nth is None` 时返回 `np.ndarray`；否则返回 `int`，未命中时为 `-1`。

### 示例
```python
import pandas as pd, xlgrab
df = pd.DataFrame({'name':['Alice','Bob','Alice'], 'age':[25,30,25]})

# 按列精确匹配
df.find_idx('name', 'Alice', mode='exact', axis='column', nth=1)   # 0
df.find_idx('name', 'Alice', axis='column', nth=None)               # array([0,2])

# Series 直接调用
df['name'].find_idx('Bob', mode='exact')                            # 1

# 包含与正则
df['name'].find_idx('li', mode='contains')                          # 0
df['name'].find_idx('^A', mode='regex')                             # 0
```

### 常见问题
- 找不到返回 `-1`；`nth=0` 会抛错（无效参数）。
- 按行查找时 `target` 是行索引（整数/标签），注意与按列时的列名区分。


