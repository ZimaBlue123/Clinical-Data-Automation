# Security Policy

## Supported Versions

下表列出本项目**接受安全更新**的版本范围：

| 模块系列 | 支持状态 | 备注 |
|---------|---------|------|
| 23_PDF_Threat_Analyzer | ✅ Active | 2026-06 新增，持续维护 |
| 15_PDF_XSS / 18_PDF_eCTD_Converter | ✅ Active | PDF 安全相关 |
| 01-12 基础模块 | ⚠️ Best effort | 关键 bug 修复；新功能按需 |
| 13-14 / 16-22 PDF 模块 | ⚠️ Best effort | 关键 bug 修复 |
| 24-33 工具类模块 | ⚠️ Best effort | 由原作者维护 |

> 模块独立版本号；本文件以模块为粒度而非仓库整体。

## Reporting a Vulnerability

**如果发现安全漏洞**（例如 PDF 解析器误报漏报、密钥扫描器失效、API Token 误提交风险等）：

### 📧 私密披露（推荐）

- **GitHub Security Advisories**：[通过仓库的 Security tab 提报](https://github.com/ZimaBlue123/1-Clinical-Data-Automation/security/advisories/new)
- 私密 issue：仓库 Settings → Security → Advisories
- **不要**直接发公开 issue（避免漏洞被利用前公开）

### 📋 报告应包含

1. **漏洞类型**：误报 / 漏报 / 注入 / 凭据泄露 / 路径穿越 / 其他
2. **影响范围**：受影响的模块（按编号）+ 文件路径
3. **复现步骤**：最小可复现样例（输入 PDF / 命令行 / 输出）
4. **影响评估**：机密性 / 完整性 / 可用性 哪方面受影响
5. **建议修复方向**（如有）
6. **披露者信息**（可选）：name / 联系方式

### ⏱️ 响应时间

| 阶段 | 时间 |
|------|------|
| 首次确认 | 7 个工作日内 |
| 修复计划 | 14 个工作日内 |
| 修复发布 | 视复杂度，30-90 天 |
| 公开披露 | 修复发布后 30-90 天（与披露者协商） |

## Scope

### ✅ 在本项目安全范围内

- **PDF 解析 / 提取 / 清理**：12-22 系列模块
- **PDF 威胁检测 / 剥离**：15_PDF_XSS / 18_PDF_eCTD_Converter / 23_PDF_Threat_Analyzer
- **敏感信息扫描**：`scripts/check_secrets.py`（pre-commit 钩子）
- **API Token 处理**：32_SAE_Extractor（OpenAI 兼容接口调用）
- **网络下载安全**：28_Paper_Batch_Download（速率限制 / 退避重试）

### ❌ 不在本项目安全范围内

- **用户本地的 PDF 内容**：本项目不存储任何用户数据
- **第三方依赖（PyMuPDF / pypdf / openpyxl 等）的漏洞**：请向对应上游报告
- **部署 / 运行环境**：用户自行负责容器 / 系统 / 防火墙配置

## Best Practices for Users

1. **永远不要**把真实 `config.yaml` / `.env` 提交到 git（已通过 `.gitignore` 兜底）
2. 跑 `pre-commit run detect-sensitive-secrets --all-files` 做最后一次全量扫描
3. 处理可疑 PDF 时，**先在隔离环境（VM / 沙箱）**跑 23_PDF_Threat_Analyzer 扫描
4. 32_SAE_Extractor 的 `SAE_API_TOKEN` 通过环境变量设置，**不要**硬编码到脚本
5. 周期性更新依赖：`pip install --upgrade pymupdf pypdf`（关注 [CVE 公告](https://github.com/pymupdf/PyMuPDF/security/advisories)）

## Hall of Fame

感谢以下披露者对本项目安全的贡献（按披露时间排序）：

- _暂无 — 等待第一位贡献者_

## License

本项目采用 [MIT License](LICENSE.md)。本安全政策不替代任何法律协议。
