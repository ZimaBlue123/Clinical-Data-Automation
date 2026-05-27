# 脚本角色约定

各编号模块（`NN_*`）内 Python 文件按职责分为三类，便于一眼区分「该运行哪个」与「仅供导入」。

| 类型 | 命名/位置 | 是否直接运行 | 说明 |
|------|-----------|--------------|------|
| **主程序** | 模块根目录，语义化文件名（如 `replace_docx.py`、`merge_ppt.py`） | ✅ | 日常批处理入口；README 中标注为「主程序」 |
| **核心库** | `lib/` 子目录 | ❌ | 被主程序或辅助工具 `import`；勿单独 `python lib/...` |
| **辅助工具** | 模块根目录，`util_` 前缀（如 `util_check_docx.py`） | ✅ 可选 | 校验、修复、环境检查、测试等；非主流程 |

仓库级共享代码在根目录 `src/`，规则提取等跨模块能力见 `config.example.yaml`。

## 多主程序模块

部分模块提供多个并列主程序（不同场景），均在模块 `README.md` 中分别说明：

| 模块 | 主程序 |
|------|--------|
| `01_Excel_Charts` | `build_charts_xlsxwriter.py`（推荐）、`build_charts_openpyxl.py` |
| `03_PPT_Merge` | `merge_ppt.py`、`ppt_engine.py`、`csr_ppt_integrator.py` |
| `09_Word_All_Tables_to_Excel` | `word_all_tables_to_excel.py`、`word_tables_merge_to_single_excel.py` |
| `12_PDF_Batch_to_Excel` | `serology_report_pdf_to_excel.py`、`fill_adr_from_pdf.py` |
| `31_SAE_Extractor` | `cli.py` |

## 已移除的无意义文件

- 仅做转发的「兼容别名」脚本（如原 `replace_date.py`）
- 一次性调试脚本（如原 `debug_zhongzhang.py`）
- 硬编码示例 / 分析残片（如原 `fixed_word_to_excel.py`、`mapping_analysis.txt`）
- 废弃目录名 `11_Word_Date_Replace`（请使用 `11_Word_Text_Replace`）；根目录误建的空目录 `Automation/`、`Data/`、`Project/`

若删除时提示「文件正在使用」，请先关闭占用该目录的 Cursor 窗口或资源管理器后再手动删除。
