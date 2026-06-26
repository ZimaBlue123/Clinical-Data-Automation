# Changelog

本项目的所有重要变更都将记录在此文件。

格式基于 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，
本项目遵循 [Semantic Versioning](https://semver.org/lang/zh-CN/)（模块独立版本号）。

## [Unreleased]

### Changed
- **17_PDF_Title_Renamer PDF Sanitizer v6.9 → v7**（Refs: 17_PDF_Title_Renamer / plans/pdf_sanitizer_v7_optimization.md）
  - **masthead 黑名单扩展**：新增 Taylor & Francis 系（emergingmicrobesinfections / tandfonline / taylorfrancis / virulence / mabs / expertopinion / expertreviewvaccines 等 14 个）、Wiley 系（advancedscience / angewandtechemie / europeanjournalof* 等 6 个）、Springer Nature 系（scientificreports / naturebiomedicalengineering 等 5 个），出版商通用词（journalhomepage / wileyonlinelibrary / springernature / academicoup / oupcom 等）
  - **BOILERPLATE_LINE 扩展**：覆盖 www.tandfonline.com / www.wiley.com / www.springer.com / link.springer.com / taylor & francis / wiley & sons / published by ... / © <year> / ISSN: <code> / creative commons 行
  - **_is_journal_masthead_only 兜底判定放宽**：≤5 词、≤48 字符，且新增 & 出版商命名模式与 infections 后缀识别（命中 Emerging Microbes & Infections）
  - **新增 _split_on_smart_colon**：英文冒号后含 Phase / Trial / Randomized / Double-blind / Multicenter 等临床试验设计词时**保留副标题**，中文冒号或无关键设计词时仍按硬截断
  - **新增 _smart_bracket_removal**：含字母/数字的技术括号（如 (Pichia pastoris)、(COVID-19)、(n=100)）保留内容；纯符号括号（如 [10.1016/...]、(:)）整段删除
  - **& → " And " 容错**：避免 Emerging Microbes & Infections 被非法字符过滤后黏连成 MicrobesInfections
  - **罗马数字斜杠分离**：_split_roman_numerals 前移到 _smart_title_case 之前，避免 I/IIa 在大小写化时退化为 I/iia；_preserve_scientific_token 同步支持 IIa / IVb 单 token 与 I_IIa 拼接形式
  - **_scan_payload 视觉层级降级**：当 hierarchy_title（最大字号命中）事后被判定为 masthead 时，提前清空并交给 _academic_title_from_plain_text 接管，避免期刊名压制正文标题
  - **单元验证脚本**：`_verify_v7.py`（worktree 内），覆盖 masthead 黑名单 / 智能冒号 / 智能括号 / 罗马数字 / boilerplate 7 项断言，全部 PASS

### Fixed
- **17_PDF_Title_Renamer：Taylor & Francis 期刊封面误识别**
  - 症状：`Emerging_Microbes_&_Infections_ISSN-2025.pdf` → `Emerging_Microbes_Infections-2025.pdf`
  - 根因：视觉层级最大字号命中期刊名 masthead，原黑名单未含 T&F 系；副标题 `Phase I/IIa trial` 因 `:` 被截断丢弃；`(Pichia pastoris)` 整块被括号过滤器删除
  - 修复后：→ `Safety_Tolerability_and_Immunogenicity_of_a_Quadrivalent_Recombinant_Norovirus_Vaccine_Pichia_Pastoris_in_Participants_Six_Weeks_of_Age_or_Older_Phase_I_IIa_Trial-2025.pdf`

### Added
- **新增模块 23_PDF_Threat_Analyzer**：PDF 静态威胁扫描 + 工业级安全剥离
  - 关键字扫描（`/JavaScript` / `/OpenAction` / `/Launch` / `/EmbeddedFiles` 等 8 类）
  - URL 提取 + 恶意协议识别（`javascript:` / `vbscript:` / `file:` 等）
  - 风险评分（LOW / MEDIUM / HIGH 三档）
  - 工业级 sanitize：fitz 优先 → pypdf 降级 → qpdf 后置校验
  - 标准自检 `--self-check`（生成测试 PDF → 跑分析 → 验证 → 剥离后复检）
  - 日志遵循 [`docs/logging_convention.md`](docs/logging_convention.md)（`action=xxx key=value`）
  - 退出码遵循 0/1/2/130 约定

### Changed
- **模块编号顺位 +1**（PDF 类聚拢）：23_File_Translator → 24，…，32_SAE_Extractor → 33
  - 22_PDF_Duplicate_Analyzer 保持原位（紧邻新 23_PDF_Threat_Analyzer）
  - 受影响：10 个模块目录 + README.md / README_EN.md / docs/script_roles.md / requirements.txt
- **23_PDF_Threat_Analyzer 性能优化**：
  - 引入 `mmap` 零拷贝扫描（≥ 1MB 文件自动启用）
  - 模块级预编译正则（`COMPILED_RISK_PATTERNS`）
  - 实测：命中数与 read() 路径完全一致；速度基本持平（Windows 上 mmap 略慢 4-6%），但**省内存**优势明显
- **README.md / README_EN.md**：
  - 修正章节编号与目录编号的历史错位 bug
  - 新增 Python 版本章节（指向 `docs/python_version.md`）
  - 新增 mmap 性能说明

### Fixed
- **`_sanitize_with_fitz` 漏剥离 catalog 级危险字典**：原版只删页面级（注释/链接/嵌入文件），不处理 `/OpenAction` / `/AA` / `/Names.JavaScript` / `/AcroForm`；现在加 pypdf 二阶段清理。**修复后剥离效果：HIGH score=27 → LOW score=0**
- **`build_test_pdf_with_javascript` 手工伪 PDF**：原版 xref 错位，fitz 无法正确解析；改用 pypdf 生成合法 PDF
- **LICENSE.md**：`32_SAE_Extractor` 引用 → `33_SAE_Extractor`（重命名时漏改）

### Documentation
- 新增 `docs/python_version.md`：Python 3.10.11 环境选择、3 种启动方式、4 种故障排查
- 新增 `CHANGELOG.md`（本文件）
- 新增 `SECURITY.md`：漏洞报告流程
- `23_PDF_Threat_Analyzer/README.md` 顶部加 Python 版本 callout

### Build & Tooling
- `requirements.txt`：`pymupdf>=1.23.0` → `>=1.27.0` · `pypdf>=4.0` → `>=6.0`
- `requirements.txt`：文件头加"完整 vs 精简 vs CI"说明
- `requirements-ci.txt`：补 `pypdf>=6.0`（23 模块 CI 测试需要）
- 新建 `requirements.lite.txt`（跳过 paddleocr / paddlepaddle，1.5GB+ 模型依赖）
- `.gitignore`：加 `__tmp_*.py`（防临时脚本污染）· 加 `.vscode/launch.json` 显式忽略（防环境变量泄露）
- `.pre-commit-config.yaml`：加 `ruff` + `ruff-format` hook（替代 flake8 + black + isort）

### Cleanup
- 删除 76 项 `__pycache__/*.cpython-314.pyc`（1.15 MB，**已 gitignore 兜底**）

### Added
- `requirements-ci.txt` 头部补充注释：定位"完整 vs 精简 vs CI"三档关系

### Changed
- **代码健壮性增强**：
  - `01_Excel_Charts/fill_clinical_table.py`：去除 5 处 f-string 无占位符冗余前缀（ruff F541 → 0）
  - `33_SAE_Extractor/sae_extractor.py`：将 `except json.JSONDecodeError: pass` 替换为 `except ... as e: logging.debug(...)`；不再静默吞异常
  - `src/excel_writer.py`：5 个公共 API 补全返回类型注解与 `Worksheet` 参数类型
- **依赖审计**：扫描全仓 81 个 .py 文件 74 个 import 声明；现有 3 个 requirements 文件覆盖完整，无未声明的第三方库
- **`.gitignore` 补全**：追加 12 个诊断/扫描输出模式（`*_baseline.txt` / `*_audit.txt` / `*_scan*.txt` 等），防止本轮及历史阶段临时文件漏入库
- **ruff 全量扫描**：0 错误（修复前 5 个 F541）
- **pytest 全量测试**：16/16 通过

### Cleanup
- 删除 14 项 `__pycache__` 目录与 `.pyc` 缓存（11_Word_Text_Replace/lib、src/、tests/ 三处）

---


---

## 历史

### 2026-05 — 模块结构稳定期
- 完成 21_PDF_Watermark_Removal / 22_PDF_Duplicate_Analyzer 等大型模块
- 全仓 Python 脚本 `compileall` 巡检通过
- 引入 `docs/script_roles.md` 与 `docs/logging_convention.md` 规范

### 早期（2024-2025）
- 01-12 模块（Excel / PowerPoint / Word / PDF 基础）逐步成型
- `15_PDF_XSS` 作为 PDF 安全清理原型
- `32_SAE_Extractor` 从外部项目迁入（[SAE-Extractor](https://github.com/...)）

### Added
- **新增模块 08_Word_Tables_to_Graphpad**：docx 临床小结 → pzfx 模板抗体数据替换（gE ↔ VZV 等）
  - `lib/docx_parser.py`：标准库 zipfile + ElementTree 解析 docx 段落 + 顶层表
  - `lib/pzfx_parser.py`：单文件 XML / ZIP 容器两种 pzfx 格式兼容
  - `lib/pzfx_writer.py`：基于字符串定位的 Subcolumn 改写（不破坏 XML 结构）
  - `lib/antibody_mapping.py`：句式识别 + Triple 解析 + 跨抗体映射
  - `poc_replicate.py`：CLI 主程序（含源值校验 + 审计 + 替换）
  - `util_probe.py`：辅助工具：同时探查 docx + pzfx 结构
  - 18 个单元测试（unittest）覆盖核心解析、改写、句式识别
  - 典型场景：gE 抗体 pzfx 模板 → VZV 抗体 pzfx；表名/列名/Subcolumn 顺序不变，仅 `<d>` 数值替换

### Changed
- **模块编号顺位调整**：原 08_Word_Tables_to_Excel + 09_Word_All_Tables_to_Excel 合并为新的 `09_Word_Tables_to_Excel`
  - 原 08_Word_Tables_to_Excel 的核心导出逻辑（`word_tables_to_excel.py`）整合到 09
  - 原 09 的两个脚本（`word_all_tables_to_excel.py` / `word_tables_merge_to_single_excel.py`）改用直接 `import` 调用，去掉 importlib hack
  - 受影响：08_Word_Tables_to_Excel/ 整目录删除、09_Word_All_Tables_to_Excel/ 改名、README.md / README_EN.md / docs/script_roles.md

