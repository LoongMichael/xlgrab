## 场景案例（仅案例描述，不含实现与代码）

### 1. 最简单固定区域（含汇总）
- 文件: D:/data/demo.xlsx
- sheet: Sheet1
- header: A2:B2
- data: A2:B6
- total: A7:B7（包含在结果末尾）
- 期望：列数=2；行数=6；列名取自 A2:B2 底行

### 2. 固定区域（不含汇总）
- 文件: D:/data/demo.xlsx
- sheet: Sheet1
- header: A9:C9
- data: A10:C16
- total: 无
- 期望：列数=3；行数=7；列名取自 A9:C9 底行

### 3. 同一 sheet 的多块（共享与不同表头混合）
- 文件: D:/data/demo.xlsx
- sheet: Sheet1
- 区域1:
  - header: A2:B2
  - data: A2:B6
  - total: A7:B7
- 区域2:
  - header: A9:C9
  - data: A10:C16
  - total: 无
- 期望：得到两个独立的结果集；区域1(2列/6行)，区域2(3列/7行)

### 4. 多块同表头，按顺序拼接
- 文件: D:/data/multi.xlsx
- sheet: Sheet1
- 共享 header: A2:B2
- 块1 data: A2:B6；total: A7:B7
- 块2 data: A10:B12；total: A13:B13
- 期望：单个结果集；列数=2；行数=（5+1）+（3+1）=10；拼接顺序为块1后块2

### 5. 固定起点 + 关键词终止（剔除表尾）
- 文件: D:/data/kw.xlsx
- sheet: Sheet1
- header: A2:B2
- data 起点: A3
- 终止关键词: “合计” 所在行的上一行为结束（offset=-1）
- total: 无
- 期望：列名取 A2:B2；行数 = 从 A3 至“合计”上一行；不包含“合计”行

### 6. 锚点起点 + 关键词终止
- 文件: D:/data/anchor.xlsx
- sheet: Sheet1
- header: A2:B2
- 锚点：在 A 列查找“名称”的第 3 次出现；数据从锚点下一行开始（row_delta=+1）
- 终止关键词: “合计” 的上一行（offset=-1）
- total: 无
- 期望：当不足 3 次时，该块跳过并记录错误；否则提取从第3次出现下一行开始至“合计”上一行

### 7. 不连续片段（显式多个固定 data）
- 文件: D:/data/discrete.xlsx
- sheet: Sheet1
- header: A2:B2
- 片段:
  - data: A2:B6
  - data: A10:B12
  - data: A15:B16
- total: 无
- 期望：单个结果集；列数=2；行数=5+3+2=10；按声明顺序拼接（不依据表内位置）

### 8. 多行表头扁平化
- 文件: D:/data/header.xlsx
- sheet: Sheet1
- header: A2:B3（两行表头）
- data: A4:B10
- total: 无
- 扁平化：按“上层_下层”拼接生成列名
- 期望：列数=2；列名形如 “类别_名称”“数量_件”

### 9. 去整行空行（可选）
- 文件: D:/data/clean.xlsx
- sheet: Sheet1
- header: A2:B2
- data: A2:B20（中间有全空行）
- total: 无
- 清理：开启“整行全空”删除
- 期望：结果中不包含任何全空行；非全空行保持原顺序

### 10. 失败与错误记录（不中断）
- 文件: D:/data/error.xlsx
- sheet: SheetNotExists（不存在）
- 任意 header/data
- 期望：返回错误项（SHEET_NOT_FOUND），但整体流程不中断

### 11. 同一文件的多 sheet 提取
- 文件: D:/data/multi_sheets.xlsx
- 区域A：sheet=Sheet1, header=A2:B2, data=A2:B10
- 区域B：sheet=Sheet2, header=A2:C2, data=A3:C8
- 期望：分别返回两个结果（或两组坐标），供后续分别入库/合并

### 12. 大表（小内存）策略
- 文件: D:/data/large.xlsx（行数较多，但本方案整体读入内存可接受）
- 读取：一次性读入目标 sheet 为 DataFrame
- 区域：固定 data 和可选 total
- 期望：确保在机器内存可承受范围内，整表读入后按坐标切片；若需更省内存，后续支持按列/按块读取策略


