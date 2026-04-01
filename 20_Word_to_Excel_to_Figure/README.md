# 20_Word_to_Excel_to_Figure

把 `input/` 下的 Word（`.doc/.docx/.rtf`）表格数据抽取出来，填入 `Template/` 骨架 Excel 的图表数据区间（`chart.series.cat/val`）并生成高保真复刻 Excel（尽量最大化保留图表/透视/OLAP 结构）。

仓库级目录说明与依赖分组见根目录 [`README.md`](../README.md)（「整体架构」、`requirements.txt`）。

## 运行环境

- Windows + 安装 Microsoft Word（用于 `pywin32` 的 Word COM）
- Python 3.8+

## 输入要求（只放原始数据）

把原始 Word/RTF 文件放到：

- `20_Word_to_Excel_to_Figure/input/`

规则：

- 后缀：`.doc / .docx / .rtf`
- 不限制文件数量；脚本会遍历目录里所有 Word/RTF
- 文件名建议包含 `part1 / part2 / part3`（如有），便于按顺序优先匹配
- 不要放 Word/WPS 打开的临时锁文件：`~$*.rtf / ~$*.docx`（脚本会跳过，但尽量保持干净）

## 骨架模板（Template）

- 骨架 Excel 放在：`20_Word_to_Excel_to_Figure/Template/`
- 要求至少包含图表的 `chart.series.cat/val` 引用区间（用于定位需要写入的数据区间）
- 通常不需要你每次提供模板；可选参数 `--template-xlsx` 在你确实有多套骨架时指定骨架路径。

## 输出

- 默认输出到：`20_Word_to_Excel_to_Figure/output/replicate_<骨架xlsx名>.xlsx`
- 如果使用 `--plan-only`：会在 `output/` 生成候选映射计划 JSON（例如 `table_mapping_plan_*.json`），供你确认“每个 Excel 子表的数据抓取来源”。

## 命令行

### 1) 生成候选映射计划（plan only）

```bash
python "20_Word_to_Excel_to_Figure/word_to_excel_to_figure.py" --input-dir "20_Word_to_Excel_to_Figure/input" --plan-only
```

可指定 plan JSON 输出路径：

```bash
python "20_Word_to_Excel_to_Figure/word_to_excel_to_figure.py" ^
  --input-dir "20_Word_to_Excel_to_Figure/input" ^
  --plan-only ^
  --plan-out-json "20_Word_to_Excel_to_Figure/output/table_mapping_plan_current.json"
```

### 2) 用你确认后的映射生成 output

```bash
python "20_Word_to_Excel_to_Figure/word_to_excel_to_figure.py" ^
  --input-dir "20_Word_to_Excel_to_Figure/input" ^
  --table-map-json "20_Word_to_Excel_to_Figure/output/table_mapping_plan_current.json"
```

可选：显示更详细日志，加 `--verbose`。

## 如何确认 table 映射（关键流程）

1. 打开 `table_mapping_plan_*.json`
2. JSON 里每个 `subtable_id` 下都有候选列表（候选项通常包含 `word_file / table_index / score`）
3. 你需要把“真正对应的候选项”设置为 `selected: true`
4. 程序会在正式生成 output 时，按 `selected` 的 `word_file + table_index` 作为该子表的数据抓取来源（未标记时默认选择候选列表第一个）

映射候选生成与 JSON 加载的**唯一实现**在：`table_mapping_logic.py`（由 `word_to_excel_to_figure.py` 导入，请勿维护两套逻辑）。

## 自检（覆盖率）

由于新输入数据的 `val` 数值可能与骨架不一致，自检不再要求“逐项数值完全相等”。

- 每个图表 series 统计 `expected_count` 与 `extracted_count`
- 若 `extracted_count / expected_count` 低于阈值，会抛出异常并停止（提示抽取覆盖率过低）

