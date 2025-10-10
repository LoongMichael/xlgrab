# xlgrab 使用手册

本手册按功能分章，配合 `README.md` 的快速开始，系统展示每个方法的语义、参数、返回值、注意事项与示例代码。

## 功能分类

### 数据查找
- **find_idx**：在列/行中查找位置（exact/contains/regex、nth）

### Excel 范围操作
- **excel_range**：从 Excel 风格区间抽取数据，可多区域纵向合并
- **offset_range**：以 1 基行列 + 偏移/裁剪选择范围
- **select_range**：区间 DSL，统一表达混合端点与偏移

### 数据处理
- **apply_header**：按 pandas read_csv 语义将上方行合并为表头

### Excel 文件操作
- **unmerge_excel**：解开Excel合并单元格并填充值
- **read_excel**：读取Excel文件的指定范围数据

## 文档索引

请按以下文档查看详细用法：

### 核心功能
- [find_idx.md](./find_idx.md) - 数据查找功能
- [excel_range.md](./excel_range.md) - Excel区间操作
- [offset_range.md](./offset_range.md) - 偏移范围选择
- [select_range.md](./select_range.md) - DSL风格区间选择
- [apply_header.md](./apply_header.md) - 表头处理

### Excel 文件操作
- [unmerge_excel.md](./unmerge_excel.md) - 合并单元格处理
- [read_excel.md](./read_excel.md) - 范围读取

## 架构说明

xlgrab 采用模块化架构设计：

```
xlgrab/
├── core.py              # 核心类（轻量委托）
├── excel/               # Excel专用功能
│   ├── merger.py       # 合并单元格处理
│   ├── reader.py       # 范围读取
│   └── range.py        # 范围操作
└── data/                # 数据操作
    ├── search.py       # 查找功能
    └── header.py       # 表头处理
```

这种设计使得功能分类清晰，易于维护和扩展。

## 快速开始

```python
import pandas as pd
import xlgrab

# 创建测试数据
df = pd.DataFrame({
    'name': ['Alice', 'Bob', 'Charlie'],
    'age': [25, 30, 35],
    'salary': [50000, 60000, 70000]
})

# 查找数据
position = df.xl.find_idx('name', 'Alice')  # 返回 0

# Excel 区间操作
data = df.xl.excel_range('A1:C3', header=True)

# 表头处理
df_with_header = df.xl.apply_header(0)

# Excel 文件操作
xlgrab.unmerge_excel("input.xlsx", "output.xlsx")
df_range = xlgrab.read_excel("data.xlsx", ranges="A1:C10")
```

更多详细用法请查看各功能的具体文档。