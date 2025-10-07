## xlgrab

一个基于 Facade 模式的 pandas 增强库。导入后，DataFrame/Series 将自动获得一组易用、贴近 Excel 思维的辅助方法，专注于“快速定位与提取数据区域”。

### 亮点
- **一行导入，方法即刻可用**：`import xlgrab` 后，`pd.DataFrame`/`pd.Series` 直接获得增强方法
- **查找定位**：`find_idx` 支持 exact/contains/regex、nth 指定、返回单个或全部命中
- **Excel 区间**：`excel_range('B2:D6', ...)`，可一次传多个区间并纵向合并
- **偏移选择**：`offset_range`/`select_range` 支持统一/分别偏移与边界裁剪
- **表头处理**：`apply_header` 与 pandas read_csv 语义一致（支持 int、list[int]、list[str]/Series、DataFrame）

## 安装

### 1) 直接从 GitHub 安装（推荐）

无需克隆代码，使用 pip 直接从仓库安装：

```bash
pip install git+https://github.com/LoongMichael/xlgrab.git
```

如需升级到仓库的最新提交：

```bash
pip install --upgrade --no-cache-dir git+https://github.com/LoongMichael/xlgrab.git
```

使用 SSH（你已配置 GitHub SSH key）：

```bash
pip install git+ssh://git@github.com/LoongMichael/xlgrab.git
```

固定到某个提交/分支/标签（示例为分支 main 与某个提交）：

```bash
pip install git+https://github.com/LoongMichael/xlgrab.git@main
pip install git+https://github.com/LoongMichael/xlgrab.git@<commit_sha>
```

### 2) 克隆源码本地开发安装（editable）

```bash
git clone https://github.com/LoongMichael/xlgrab.git
cd xlgrab
pip install -e .
```

如需安装测试依赖，请使用 `requirements.txt`（仅当你打算本地运行用例）：

```bash
pip install -r requirements.txt
```

### 3) 运行环境与依赖

- Python: >= 3.7（参考 `setup.py` 的 `python_requires`）
- 运行依赖：
  - pandas >= 1.3.0
  - numpy >= 1.20.0
  - openpyxl >= 3.0.0（仅在使用 `excel_range` 时需要）

可使用镜像源（示例以中国大陆环境）：

```bash
pip install -i https://pypi.tuna.tsinghua.edu.cn/simple git+https://github.com/LoongMichael/xlgrab.git
```

### 4) 安装后快速验证

```python
import pandas as pd
import xlgrab  # 导入后会为 pandas 注册扩展方法

df = pd.DataFrame({"name": ["Alice", "Bob"], "age": [25, 30]})

# 若安装成功，下行方法应可调用且无 AttributeError
_ = df.find_idx('name', 'Alice', mode='exact', axis='column', nth=1)
print('xlgrab OK')
```

## 快速开始

```python
import pandas as pd
import xlgrab  # 导入后自动为 pandas 注册扩展方法

df = pd.DataFrame({
    'name': ['Alice', 'Bob', 'Charlie', 'David', 'Eve'],
    'age': [25, 30, 35, 28, 32],
    'salary': [50000, 60000, 70000, 55000, 65000],
    'dept': ['IT', 'HR', 'IT', 'Finance', 'Marketing']
})

# 查找位置（按列）
df.find_idx('name', 'Alice', mode='exact', axis='column', nth=1)     # 0
df.find_idx('dept', 'IT', mode='exact', axis='column', nth=None)      # array([0, 2])

# 查找位置（按行）
df.find_idx(0, '^A', mode='regex', axis='row', nth=1)                 # 0

# Excel 区间到 DataFrame，支持多区域合并
df.excel_range('B2:D4', 'F10:H12', header=True)

# 偏移选择（以 1 基行列坐标）
df.offset_range(1, 5, 2, 6, offset_rows=1, offset_cols=-1, clip_to_bounds=True)

# DSL 选择 + 偏移（先解析，再调用 offset_range 执行切片）
df.select_range(start='A2', end=('cell','C5'), offset_rows=1, offset_cols=0, clip=True)

# 表头设置（read_csv 语义）
df.apply_header(True)            # True 等价 0 行作为表头
df.apply_header(0)               # 第 0 行为表头
df.apply_header([0, 1])          # 多行表头；header_join=None 则生成 MultiIndex
df.apply_header(['客户', '金额', '日期'])  # 直接用给定列表命名（自动规范化、递增去重）
```

## 核心 API

### find_idx(target, q, mode='exact', na=False, nth=1, axis='column')
- **target**: 列名或行索引
- **mode**: exact | contains | regex
- **nth**: None 返回全部；>0 第 n 次；<0 倒数第 n 次
- **axis**: 'column' | 'row'

### excel_range(*ranges, header=True, index_col=None)
- 支持一次传入多个 Excel 区间字符串，自动按行合并
- `header=True` 将首行作为列名；`index_col` 支持列名或位置
- 使用 `openpyxl.utils.coordinate_to_tuple` 解析坐标

### offset_range(start_row, end_row, start_col, end_col, ...)
- 统一偏移：`offset_rows`、`offset_cols`
- 分别偏移：`offset_start_row/end_row/start_col/end_col`
- `clip_to_bounds=True` 自动裁剪；否则越界报错

### select_range(start 或 start_row/col/end_row/col, clip=True, ...offsets)

`select_range` 提供一个表达能力强、贴近 Excel 与查找语义的“区间 DSL”。它将多种端点描述转换为最终 `iloc` 切片，并在末尾统一复用 `offset_range` 执行偏移与边界处理。

- 支持的端点参数（四个边界可独立提供，缺省时有默认值）：
  - `start`: 一次性指定起点（可为单元格、仅行或仅列、或 find 规范）
  - `start_row`, `end_row`, `start_col`, `end_col`: 覆盖对应维度

- 端点 DSL 说明（均区分“行语境”与“列语境”）：
  - 字符串
    - 'A2'：单元格（同时指定行与列）
    - 'F' / 'AA'：列（列语境）
  - 整数（Excel 习惯的 1 基）：例如 3 表示第 3 行/列
  - 元组/列表
    - ('cell', 'A2')：显式单元格
    - ('row', 10)：显式行
    - ('col', 'F' | 6)：显式列
    - ('find-row', target, q, {mode, nth, na, flags})：按列搜索“行边界”
    - ('find-col', target, q, {mode, nth, na, flags})：按行搜索“列边界”

- find 规范与 `find_idx` 一致：
  - `mode`: exact | contains | regex
  - `nth`: None 返回全部；>0 第 n 次；<0 倒数第 n 次
  - `na`, `flags`：传递给底层 `str.contains`/正则
  - `target`：在“行边界”场景下是列名/索引；“列边界”场景下是行索引/标签

- 边界默认值与顺序规范：
  - 未指定时默认 `start_row=1`, `start_col=1`, `end_row=末行`, `end_col=末列`
  - 起止顺序会自动校正（若 start>end 会交换）

- 偏移与裁剪：
  - 统一偏移：`offset_rows`, `offset_cols`
  - 分别偏移：`offset_start_row`, `offset_end_row`, `offset_start_col`, `offset_end_col`
  - `clip=True` 将越界自动裁剪；`clip=False` 越界将抛出错误
  - 实际偏移与切片由 `offset_range` 执行，保证行为一致

- 常见用法示例：
```python
# 1) 起止用单元格
df.select_range(start='A2', end=('cell','C5'))

# 2) 起始用行，列为默认（终止默认到底）
df.select_range(start_row=('row', 10))

# 3) 列用 Excel 列字母，行用整数（1 基）
df.select_range(start_col='B', end_col='F', start_row=2, end_row=20)

# 4) 用 find 指定 4 个边界（可独立配置 mode/nth/na/flags）
df.select_range(
    start_row=('find-row', 'name', 'Alice', {'mode': 'exact', 'nth': 1}),
    end_row=('find-row', 'name', 'Eve', {'mode': 'exact', 'nth': 1}),
    start_col=('find-col', 0, 'age', {'mode': 'exact'}),
    end_col=('find-col', 0, 'salary', {'mode': 'exact'})
)

# 5) 偏移：统一偏移行+1列-1，并自动裁剪
df.select_range(start='A2', end=('cell','C5'), offset_rows=1, offset_cols=-1, clip=True)

# 6) 偏移：分别偏移（起始行+1，结束行+2，起始列-1，结束列不变）
df.select_range(
    start='A2', end=('cell','C5'),
    offset_start_row=1, offset_end_row=2,
    offset_start_col=-1, offset_end_col=0,
)
```

### apply_header(header, header_join="_")
- 与 pandas read_csv 语义保持一致：
  - True/0/1...：使用指定“0 基”行作为表头
  - [i, j, ...]：多行表头；`header_join=None` 生成 MultiIndex，否则用分隔符合并
  - list[str]/tuple/Series：直接作为列名（自动规范化、重复名递增 `_1/_2/...`）
  - DataFrame：外部多行表头来源

规范化规则：替换常见特殊字符为下划线；合并连续下划线；去除首尾下划线。

## 测试

```bash
python -m unittest tests/test_apply_header.py -v
```

或运行全部用例：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

## 依赖

- pandas >= 1.3.0
- numpy >= 1.20.0
- openpyxl >= 3.0.0（使用 excel_range 时需要）

## 许可证

MIT License

## 文档

更多使用说明与示例请查看 `docs/manual/`：

- `docs/manual/index.md`：文档索引
- `docs/manual/find_idx.md`：`find_idx` 用法与示例
- `docs/manual/excel_range.md`：`excel_range` 用法与示例
- `docs/manual/offset_range.md`：`offset_range` 用法与示例
- `docs/manual/select_range.md`：`select_range` 用法与示例
- `docs/manual/apply_header.md`：`apply_header` 用法与示例
