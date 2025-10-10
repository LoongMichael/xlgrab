# apply_header - 表头处理

## 功能说明

`apply_header` 用于将 DataFrame 顶部若干行作为表头，规则与 pandas read_csv 语义一致。提供列名清洗与重复名递增功能，支持多种表头格式。

## 函数签名

```python
df.xl.apply_header(header=True, header_join="_", inplace=False)
```

### 参数
- header：
  - True/0：第 0 行作为表头；
  - int：第 N 行作为表头；
  - list[int]：多行表头；
  - list[str]/tuple/Series：直接作为列名；
  - DataFrame：外部多行表头来源。
- header_join：多行表头时用于合并为单层列名的分隔符；None 则生成 MultiIndex。

### 返回值
- 处理后的新 DataFrame（不修改原始 df）。

### 示例
```python
import pandas as pd, xlgrab
df = pd.DataFrame({
  'A':['姓名','Alice','Bob'], 'B':['年龄',25,30], 'C':['部门','IT','HR']
})

# 单行表头（等价 True → 0 行）
df.apply_header(0)

# 多行表头合并为单层
hdr = pd.DataFrame({
  0:['一级','姓名','年龄','部门'],
  1:['二级','中文','数字','中文']
})
df.apply_header(hdr, header_join='_')

# 生成 MultiIndex 表头
df.apply_header(hdr, header_join=None)

# 直接指定列名列表（会清洗与去重）
df.apply_header(['姓名','年龄','部门'])
```

### 注意事项
清洗规则会替换常见特殊字符为下划线，合并连续下划线，并去除首尾下划线；多行表头合并后数据部分会重建 0..N 索引。


