# 11. Word 文档批量文本替换（`11_Word_Text_Replace`）

对 `.docx` 做 **OOXML 级批量查找替换**（不依赖 Word 进程），适用于 CSR/方案中的日期占位符、研究编号、固定短语等。

覆盖范围：正文、表格、页眉页脚；支持 **跨 run 拆分** 与修订删除文本（`w:delText`）。

脚本角色约定见仓库 [`docs/script_roles.md`](../docs/script_roles.md)。

## 内置规则（可扩展）

| 类型 | 查找示例 | 替换为 |
|------|----------|--------|
| 日期（斜杠） | `2026/XX/XX` | `2026/05/27` |
| 日期（中文） | `2026年XX月XX日` | `2026年05月27日` |
| 研究编号 | `YDSWX（TVAX-009）-004（Ⅳ）` 等 | `YDSWX（TVAX-009）-004（III）` |

新增规则：编辑 `lib/ooxml_replace.py` 中 `build_*_rules()`，或后续在 `replace_rules.yaml` 中维护（见 `replace_rules.example.yaml`）。

## 目录

| 目录 | 说明 |
|------|------|
| `input/` | 待处理 `.docx`（不提交，见根目录 `.gitignore`） |
| `output/` | 输出 `*_updated.docx`（不提交） |

## 快速开始

```bash
cd 11_Word_Text_Replace
# 将 docx 放入 input/
python replace_docx.py --yes
```

按规则类型：

```bash
python replace_docx.py --only-study-id --yes   # 仅研究编号
python replace_docx.py --only-dates --yes      # 仅日期占位符
```

## 校验（辅助工具）

```bash
python util_check_docx.py --latest
python util_check_docx.py --latest --mode study
```

## 文件说明

| 文件 | 角色 |
|------|------|
| `replace_docx.py` | **主程序** |
| `lib/ooxml_replace.py` | 核心库（OOXML 替换与规则） |
| `util_check_docx.py` | 辅助：输出校验 |
| `replace_rules.example.yaml` | 自定义规则示例（预留） |

## 依赖

- `lxml`（见仓库根目录 `requirements.txt`）
