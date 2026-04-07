# Template（骨架 Excel，敏感文件不提交）

本目录用于存放“骨架 xlsx”，用作 `word_to_excel_to_figure.py` 的图表结构与 `chart.series(cat/val)` 引用区间定位载体。

注意：
- 本目录下的 `*.xlsx` 通常为敏感/体积较大文件，不应提交到 Git（已在根目录 `.gitignore` 中忽略）。
- 你可以在本地放入你之前成功且不触发 Office 修复弹窗的骨架 xlsx。
- 若需要指定骨架，可运行：`word_to_excel_to_figure.py --template-xlsx "path/to/skeleton.xlsx"`。

# Template（骨架 Excel）

该目录用于放置“骨架 xlsx”，用来保持 `chart.series.cat/val` 引用区间、图表/透视/OLAP 等复杂结构。

重要说明：
- 本目录里的 `.xlsx` 骨架文件默认不会提交到 Git（已在根目录 `.gitignore` 中忽略）。
- 你需要在本地准备好骨架文件（至少 1 个），脚本会优先使用目录内“最佳匹配”的骨架；如有多套骨架，也可以通过 `--template-xlsx` 指定。

脚本入口：`word_to_excel_to_figure.py`

