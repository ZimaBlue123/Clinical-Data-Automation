# 01_Excel_Charts - Excel图表数据处理

本模块用于处理临床试验中的Excel数据填充和图表生成。

## 目录结构

```
01_Excel_Charts/
├── input/                       # 输入文件目录（原始数据）
├── output/                      # 输出文件目录（处理结果）
├── fill_clinical_table.py       # 临床表格数据填充（统一入口：GMC + GMI + 阳转率）
├── apply_template_charts.py
├── build_charts_openpyxl.py
└── build_charts_xlsxwriter.py
```

## 脚本说明

### fill_clinical_table.py - 临床表格数据填充（统一入口）

统一处理GMC（几何平均浓度）、GMI（几何平均倍数）和阳转率三类临床试验数据表格的自动填充。

**支持表格类型：**
- `gmc`：GMC（几何平均浓度）表格
- `gmi`：GMI（几何平均倍数）表格
- `yangzhuai`：阳转率表格
- `all`：自动检测所有支持的子表（默认）

**用法：**

```bash
# 处理所有支持的子表（自动检测GMC/GMI/阳转率）
python fill_clinical_table.py input/TVAX-006.xlsx

# 仅处理GMC子表
python fill_clinical_table.py input/TVAX-006.xlsx --type gmc

# 仅处理GMI子表
python fill_clinical_table.py input/TVAX-006.xlsx --type gmi

# 仅处理阳转率子表
python fill_clinical_table.py input/TVAX-006.xlsx --type yangzhuai

# 指定具体子表
python fill_clinical_table.py input/TVAX-006.xlsx --sheets "总体GMC,40-59岁GMC"

# 自定义输出目录
python fill_clinical_table.py input/TVAX-006.xlsx --output-dir ./output

# 显示详细日志
python fill_clinical_table.py input/TVAX-006.xlsx -v
```

**参数：**
- `excel_file`：Excel文件路径（必填）
- `--type/-t {gmc,gmi,yangzhuai,all}`：表格类型（默认：all）
- `--sheets/-s`：指定子表（逗号分隔，优先级高于--type）
- `--output-dir/-o`：输出目录（默认：同级output目录）
- `--output-name/-n`：输出文件名（默认：输入文件名）
- `--no-include-pre`：GMC表格不填充免前（第3行）
- `--verbose/-v`：显示详细日志

**GMC数据格式：**
- GMC表格第1-7行：组别标题、子标题、免前 + 4个时间点
- 源数据：
  - 行12（GMC 95%CI）→ 免前
  - 行17/22/27/32（LS GMC 95%CI）→ 4个时间点
- 源数据格式：`"768.17(507.87, 1161.89)"` 或 `"644.46 (280.78, 1479.20)"`

**GMI数据格式：**
- GMI表格第1-7行：组别标题、子标题、免前 + 4个时间点
- 源数据：脚本动态扫描定位"GMI (95%CI)"行（适配不同子表行号差异）
  - "60岁以上GMI"使用51行结构（默认52行）
  - 默认源行：18/29/41/52（脚本自动适配）
- 源数据格式：`"1.12 (0.91, 1.38)"` 与GMC一致

**阳转率数据格式：**
- 阳转率表格第1-7行：组别标题、子标题、免前（不填）+ 4个时间点
- 源数据：脚本自动扫描定位
  - 标题行：第1列含"一免后"或"全免后"
  - "阳转例数（阳转率）"行：格式 `"24 (75.00)"`
  - "95%CI"行：格式 `"56.60, 88.54"`
- 自动适配"总体阳转率"等含"年龄组"行的复杂结构

**列结构（5个组 × 3列）：**
| 列范围 | 组别 |
|--------|------|
| B-D | 低剂量佐剂组（均值、上限、下限） |
| E-G | 高剂量佐剂组 |
| H-J | 低剂量试验组 |
| K-M | 高剂量试验组 |
| N-P | 安慰剂组 |

**适用场景：**
- 临床试验体液免疫数据中GMC/GMI/阳转率表格的自动填充
- 支持任意子表名（中文、含特殊字符）
- 自动识别源数据位置，无需手动指定行号
- 多子表批量处理
- 动态适配不同子表的行号结构

## 其他脚本

### apply_template_charts.py
应用图表模板生成图表。

### build_charts_openpyxl.py / build_charts_xlsxwriter.py
使用不同库生成Excel图表。

## 输入文件示例

典型输入文件：`TVAX-006 Phase 1_体液免疫VZV.xlsx`

包含多个子表：
- 总体GMC、40-59岁GMC、60岁以上GMC
- 总体GMI、40-59岁GMI、60岁以上GMI
- 总体阳转率、40-59岁阳转率、60岁以上阳转率

## 变更日志 (Changelog)

### [1.0.0] - 2026-06-25

#### 新增 (Added)
- **`fill_clinical_table.py`**：临床试验表格数据填充的统一入口脚本
  - 支持 `GMC`（几何平均浓度）、`GMI`（几何平均倍数）和 `阳转率`（seroconversion rate）三类表格
  - 通过 `detect_sheet_type()` 自动识别子表类型，无需手动指定
  - 动态扫描源数据行号，适配不同子表结构
  - 完整类型提示、logging、异常捕获、边界条件处理
  - 适配"60岁以上GMI"的51行特殊结构（默认52行）

#### 变更 (Changed)
- 整合 `fill_gmc_table.py` 和 `fill_yangzhuai_table.py` 为单一入口
- 移除 `__pycache__` 目录及 `.pyc` 缓存文件
- 清理调试脚本残留
