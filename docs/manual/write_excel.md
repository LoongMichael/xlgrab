# 写入 Excel（writer）

提供向现有 Excel 文件高效写入数据的功能。对外暴露三个核心函数：
- `to_sheet_many`: **(推荐)** 自动分批写入多个任务，性能最高。
- `write_to_excel`: 写入单个DataFrame。
- `write_range_to_excel`: 写入二维列表或元组。

> 注意：本模块专注“修改已有 Excel 文件”，内部使用 openpyxl。

## 接口概览

- `to_sheet_many(tasks)`: **（推荐）** 自动按文件名分批，高效写入多个任务。
- `write_to_excel(..., merge_policy='unmerge')`: 写入单个 DataFrame，提供完整参数控制。
- `write_range_to_excel(..., merge_policy='unmerge')`: 写入二维列表/元组的简化函数。

## 批量写入（推荐）

`to_sheet_many` 是执行所有批量写入任务的**首选方法**。它最简单、最高效。

你只需要准备一个任务列表，函数会自动按文件名分组，并对每个文件执行一次“打开-写入-保存”操作。

```python
from xlgrab.excel.writer import to_sheet_many
import pandas as pd

# 假设 df1, df2, df3 已定义
df1 = pd.DataFrame({'A': [1]})
df2 = pd.DataFrame({'B': [2]})
df3 = pd.DataFrame({'C': [3]})

# 1. 定义所有写入任务（可混合不同文件）
all_tasks = [
    # 写入 file1.xlsx
    {"excel_name": "file1.xlsx", "df": df1, "sheet_name": "Report"},
    {"excel_name": "file1.xlsx", "df": df2, "sheet_name": "RawData", "start_row": 10},

    # 写入 file2.xlsx
    {"excel_name": "file2.xlsx", "df": df3, "sheet_name": "Summary"},
]

# 2. 一次调用，全部完成
to_sheet_many(all_tasks)
```

## 单次写入

如果你只需要写入单个 DataFrame，可以使用 `write_to_excel`。

```python
from xlgrab.excel.writer import write_to_excel

df = pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6]})

# 将 df 写到 Sheet1 的 B2 开始（含列名）
write_to_excel(df, "test.xlsx", sheet_name="Sheet1", start_row=2, start_col=2, header=True)
```

## 参数说明（精简）

- `excel_name`: 目标 Excel 文件路径（若不存在，将创建新文件）。
- `sheet_name`: 工作表名或索引（0 表示第一个工作表）。
- `start_row`/`start_col`: 起始行/列，均从 1 开始计数。
- `header`/`index`: 是否写入列名/行索引（仅 `write_to_excel` 支持）。
- `merge_policy`: 合并单元格处理策略 (默认为 `'unmerge'`)。
  - `'unmerge'`: 在写入前自动取消与目标区域重叠的合并单元格（推荐）。
  - `'error'`: 如果与合并单元格冲突，则抛出 `ValueError`。

## 使用建议与注意事项

- **性能**：对于所有批量写入场景，请使用 `to_sheet_many` 以获得最佳性能。
- **合并单元格**：默认情况下，写入函数会自动取消重叠的合并单元格以避免报错。您可以通过设置 `merge_policy='error'` 来禁用此行为。
- **起始坐标**：所有坐标均从 1 开始计数，例如 B2 对应 `start_row=2, start_col=2`。

## 异常与提示

- `ValueError`: 参数类型、起止行列非法或与合并单元格冲突 (`merge_policy='error'`) 时抛出。
- `UserWarning`: 当 DataFrame 尺寸大于目标区域导致数据被截断时，会给出提醒。
