# xlgrab - 极简Excel数据提取库

> 专注简洁，专注实用 - 让Excel数据提取变得简单

## 🎯 设计理念

- **极简API** - 用户只需关心"提取什么数据"
- **函数式设计** - 纯函数，无状态，易测试
- **单一职责** - 每个函数只做一件事
- **渐进式复杂度** - 从简单到复杂，按需使用

## 🚀 快速开始

### 安装

```bash
pip install xlgrab
```

### 基本用法

```python
import xlgrab

# 1. 简单区域提取
result = xlgrab.extract_simple("data.xlsx", "Sheet1", "A1:C10")
print(result.data)

# 2. 带表头提取
result = xlgrab.extract_with_header("data.xlsx", "Sheet1", "A1:C1", "A2:C10")
print(result.columns)  # 列名
print(result.data)     # 数据

# 3. 自动表格提取
result = xlgrab.extract_table("data.xlsx", "Sheet1", "A1")
df = result.to_dataframe()  # 转换为pandas DataFrame
```

## 📚 API 参考

### 核心函数

#### `extract_simple(file_path, sheet, area)`
提取固定区域数据

```python
# 提取A1:C10区域
result = xlgrab.extract_simple("data.xlsx", "Sheet1", "A1:C10")
```

#### `extract_with_header(file_path, sheet, header_area, data_area)`
提取带表头的数据

```python
# 表头在A1:C1，数据在A2:C10
result = xlgrab.extract_with_header("data.xlsx", "Sheet1", "A1:C1", "A2:C10")
```

#### `extract_table(file_path, sheet, start_cell)`
自动检测表格边界

```python
# 从A1开始自动检测表格
result = xlgrab.extract_table("data.xlsx", "Sheet1", "A1")
```

#### `extract_list(file_path, sheet, column, start_row)`
提取单列列表数据

```python
# 提取A列从第2行开始的数据
result = xlgrab.extract_list("data.xlsx", "Sheet1", "A", 2)
```

### 高级用法

#### 锚点提取
通过文本查找区域

```python
# 在A列找"姓名"，向下偏移1行
spec = xlgrab.anchor_spec("Sheet1", "A", "姓名", 1, (1, 0))
result = xlgrab.extract("data.xlsx", [spec])
```

#### 多区域提取
一次提取多个区域

```python
specs = [
    xlgrab.range_spec("Sheet1", "A1:C5"),
    xlgrab.range_spec("Sheet1", "A7:C10"),
    xlgrab.anchor_spec("Sheet1", "A", "总计", 1, (1, 0))
]
result = xlgrab.extract("data.xlsx", specs)
```

### 结果处理

```python
result = xlgrab.extract_simple("data.xlsx", "Sheet1", "A1:C10")

# 获取数据
data = result.data          # List[List[Any]]
columns = result.columns    # List[str]
errors = result.errors      # List[str]

# 转换为其他格式
df = result.to_dataframe()  # pandas DataFrame
dict_list = result.to_dict()  # List[Dict[str, Any]]
```

## 🔧 区域语法

### 固定区域
- `"A1:C10"` - 从A1到C10
- `"A1:last"` - 从A1到最后一行
- `"A1:lastcol"` - 从A1到最后一列
- `"A1:lastlast"` - 从A1到最后一行最后一列

### 锚点区域
```python
xlgrab.anchor_spec(sheet, column, text, occurrence, offset)
```

- `sheet`: 工作表名称
- `column`: 搜索列（如"A", "B"）
- `text`: 要查找的文本
- `occurrence`: 第几次出现（默认1）
- `offset`: 偏移量 (行偏移, 列偏移)

## 📝 使用示例

### 示例1: 提取员工信息表

```python
import xlgrab

# 提取员工信息（表头+数据）
result = xlgrab.extract_with_header(
    "employees.xlsx", 
    "Sheet1", 
    "A1:D1",  # 表头：姓名、年龄、部门、工资
    "A2:D100" # 数据行
)

# 转换为DataFrame进行分析
df = result.to_dataframe()
print(df.head())
print(df.describe())
```

### 示例2: 提取多个报表

```python
# 提取多个报表区域
specs = [
    xlgrab.range_spec("Q1报表", "A1:F20"),
    xlgrab.range_spec("Q2报表", "A1:F20"),
    xlgrab.anchor_spec("Q3报表", "A", "Q3数据", 1, (1, 0))
]

result = xlgrab.extract("reports.xlsx", specs)
```

### 示例3: 提取列表数据

```python
# 提取产品名称列表
products = xlgrab.extract_list("products.xlsx", "Sheet1", "A", 2)
product_names = [row[0] for row in products.data]
```

## 🎨 设计优势

### 相比传统方式

**传统方式（复杂）:**
```python
# 需要理解多个概念：Rule, HeaderSpec, BlockSpec, AnchorSpec...
rule = Rule(
    rule_id="emp1",
    sheet_name="Sheet1", 
    header=HeaderSpec(header_range="A1:D1"),
    blocks=[BlockSpec(type="fixed", range_a1="A2:D100")]
)
result = extract_file("data.xlsx", [rule])
```

**xlgrab（极简）:**
```python
# 直接表达意图
result = xlgrab.extract_with_header("data.xlsx", "Sheet1", "A1:D1", "A2:D100")
```

### 核心优势

1. **学习成本低** - 5分钟上手
2. **代码简洁** - 减少90%的样板代码
3. **功能强大** - 支持所有常见场景
4. **易于测试** - 纯函数，无副作用
5. **向后兼容** - 保留底层函数

## 🔄 迁移指南

如果你在使用旧版本，可以这样迁移：

```python
# 旧版本
from xlgrab.api import extract_file
from xlgrab.models import Rule, HeaderSpec, BlockSpec

# 新版本
import xlgrab

# 旧版本复杂调用
rule = Rule(
    rule_id="data1",
    sheet_name="Sheet1",
    header=HeaderSpec(header_range="A1:C1"),
    blocks=[BlockSpec(type="fixed", range_a1="A2:C10")]
)
result = extract_file("data.xlsx", [rule])

# 新版本简单调用
result = xlgrab.extract_with_header("data.xlsx", "Sheet1", "A1:C1", "A2:C10")
```

## 📄 许可证

MIT License

## 🤝 贡献

欢迎提交Issue和Pull Request！

---

**xlgrab** - 让Excel数据提取变得简单 ✨