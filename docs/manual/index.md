## xlgrab 使用手册

本手册按功能分章，配合 `README.md` 的快速开始，系统展示每个方法的语义、参数、返回值、注意事项与示例代码。

- find_idx：在列/行中查找位置（exact/contains/regex、nth）
- excel_range：从 Excel 风格区间抽取数据，可多区域纵向合并
- offset_range：以 1 基行列 + 偏移/裁剪选择范围
- select_range：区间 DSL，统一表达混合端点与偏移
- apply_header：按 pandas read_csv 语义将上方行合并为表头

请按以下文档查看：

- ./find_idx.md
- ./excel_range.md
- ./offset_range.md
- ./select_range.md
- ./apply_header.md


