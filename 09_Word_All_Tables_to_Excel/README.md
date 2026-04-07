# 09_Word_All_Tables_to_Excel

把 `input/` 下的所有 Word/RTF 文件的**全部顶层表格**导出为 Excel（`.xlsx`）。

## 功能
- 遍历 `input/` 目录下的 `.doc/.docx/.rtf`
- 每个 Word 文件输出一个 Excel：`output/<word文件名>_all_tables.xlsx`
- Excel 内每个 Word 顶层表（`Document.Tables`）对应一个 sheet
- 表头支持多行合并（`--header-rows`），并尽量保持 Excel 兼容样式（冻结窗格、表头样式、边框等）

## 依赖与兼容性
- Windows + 安装 Microsoft Word（COM：`pywin32`）
- 该模块复用项目中 `08_Word_Tables_to_Excel/word_tables_to_excel.py` 的 COM 抽取与 Excel 写入逻辑，因此兼容性/健壮性与 08 保持一致

## 输入输出
- 输入：`09_Word_All_Tables_to_Excel/input/`
- 输出：`09_Word_All_Tables_to_Excel/output/`

## 快速开始
1. 把 Word 文件放入 `input/`
2. 在该目录运行：
```bash
cd 09_Word_All_Tables_to_Excel
python word_all_tables_to_excel.py
```

## 合并为单个列表（你当前需求）
若要把 `input/` 下多份 Word 的相关表格合并成**一个** Excel（列结构为 `样品ID + 五项指标(数值/说明)`），运行：

```bash
cd 09_Word_All_Tables_to_Excel
python word_tables_merge_to_single_excel.py
```

默认输出：`output/word_tables_merged.xlsx`  
可通过 `-o` 指定输出文件。缺失 `Anti-HBc / Anti-HBe / HBeAg` 的文档会自动留空。

可选：如果你要做跨模块对账（比如把 PDF 侧结果回填到 Word 侧），可以在 `word_tables_merge_to_single_excel.py` 使用 `--reference-excel <Excel.xlsx>`。

## 常用参数
- `--header-rows N`：表头占用的行数（默认 `1`；多级表头可设为 2/3）
- `--dry-run`：只读取并打印行列统计，不生成 xlsx（方便先验证表结构）
- `--quiet`：减少导出过程输出
- `--skip-existing`：若 output 已存在则跳过
- `--no-backup-existing`：覆盖前不备份旧文件（默认会备份）
- `--overwrite`：直接覆盖输出（不备份）

## 选择/高级用法（可选）
默认会导出“全部顶层表格”。如需只导出部分表，可使用：
- `--table-index N` 或 `--table-indices "1,3,5"`
- `--header-keywords "关键字1,关键字2"`
- `--merge-tables-from A --merge-tables-to B`：用于 Word 把视觉连续表拆成多个顶层表的情况（启用后会仅导出合并区间结果）

