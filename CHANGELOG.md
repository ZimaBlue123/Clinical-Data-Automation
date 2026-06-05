# Changelog

本项目的所有重要变更都将记录在此文件。

格式基于 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，
本项目遵循 [Semantic Versioning](https://semver.org/lang/zh-CN/)（模块独立版本号）。

## [Unreleased]

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
