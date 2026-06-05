# 23_PDF_Threat_Analyzer

PDF 静态威胁扫描 + 工业级安全剥离（无渲染 / 静态特征）。

> 🐍 **Python 版本要求**：本模块经过 `Python 3.10.11` 测试。
>
> 部分 Windows 环境同时存在多个 Python（如 3.10 系统安装版 + 3.14 MSYS2/UCRT64），且 `PATH` 默认指向的 Python **可能没装 pypdf / pymupdf**，跑起来会报 `ModuleNotFoundError`。
>
> 正确运行方式（任选其一）：
> ```powershell
> # 方式 A：用 py launcher 指定 3.10
> py -3.10 23_PDF_Threat_Analyzer\pdf_threat_analyzer.py --self-check
>
> # 方式 B：用 Python 3.10 完整路径
> & "C:\Users\Administrator\AppData\Local\Programs\Python\Python310\python.exe" `
>     23_PDF_Threat_Analyzer\pdf_threat_analyzer.py --self-check
> ```
>
> 详细说明见 [`docs/python_version.md`](../docs/python_version.md)。

## 适用场景

- 收到含 PDF 的可疑邮件附件，扫描后再决定是否打开
- 临床文档流转前对 PDF 做"威胁白名单"前置检查
- 第三方提供的 PDF 在归档前**剥离已知威胁对象**（JavaScript / OpenAction / Launch / 嵌入文件等）
- 自动化流水线中的一道安全闸（与 `15_PDF_XSS` 互补：本模块**优先**分析 + 报告，`15_PDF_XSS` 专注批处理清理）

> ⚠️ **本模块不会阻止 0-day 漏洞**——它按已知威胁字典做保守剥离，零日绕过、混淆 payload、政治性内容不在检测范围。交付前建议用专业病毒扫描引擎二次校验。

## 与同类型模块的关系

| 模块 | 重点 | 互补点 |
|------|------|--------|
| `15_PDF_XSS/pdf_xss_clean.py` | 批量清理 XSS / 注释 / 嵌入文件（PyMuPDF） | 命令简洁，批处理场景 |
| `18_PDF_eCTD_Converter` | eCTD 合规（6.26 字体 / 6.5-6.8 书签 + XSS） | eCTD 提交场景 |
| **`23_PDF_Threat_Analyzer`**（本模块） | **威胁分析 + 风险评分 + 工业级剥离 + 自检** | 风险报告 + 可选 sanitize，含标准自检 |

## 目录约定

| 目录 | 说明 |
|------|------|
| `input/` | 待分析 PDF（`.gitignore` 已忽略，保留 `README.md` 占位） |
| `output/` | `threat_report_*.json` + `threat_summary_*.txt` + `*_sanitized.pdf`（`.gitignore` 已忽略） |

## 依赖

| 包 | 用途 | 必需？ | 来源 |
|----|------|--------|------|
| `pypdf` ≥ 4.0 | 结构化解析 / 基础 sanitize / 自检 | **必需** | `requirements.txt` |
| `pymupdf` (fitz) ≥ 1.23 | **工业级 sanitize**（mupdf 内核级剥离） | 可选（推荐） | `requirements.txt` |
| `mmap` (标准库) | **零拷贝扫描**（≥ 1MB 文件自动启用，节省内存） | 内置 | Python 标准库 |
| `qpdf` (CLI) | 可选线性化校验 | 可选 | 系统 PATH，未装则 graceful skip |
| `mutool` (CLI) | 可选备用工具 | 可选 | 系统 PATH，未装则 graceful skip |

降级链：`fitz` → `pypdf` → 报错（无 sanitize 引擎）

## 性能优化：mmap 零拷贝扫描

`_scan_raw_binary()` 在 2026-06 commit 引入 mmap 零拷贝优化：

- **大文件（≥ 1MB）**：用 `mmap.mmap(f.fileno(), 0, ACCESS_READ)` 把文件直接映射到内存，正则直接搜索映射区，**不复制文件到用户态缓冲**。优势是 GB 级大文件不会触发 OOM。
- **小文件（< 1MB）**：用传统 `f.read()`，避免 mmap 建链开销（小文件 mmap 反而更慢）。
- **预编译正则**：`COMPILED_RISK_PATTERNS` 在模块级一次性编译所有 8 个关键字的正则模式，扫描时**零编译开销**。

实测对比（Windows 10 + Python 3.10 + 5 次最佳值）：

| 文件大小 | read() | mmap() | 备注 |
|---------|--------|--------|------|
| 0.1 MB | 0.30 ms | 0.31 ms | read 略快（小文件 mmap 建链开销占比大） |
| 1.0 MB | 2.17 ms | 2.15 ms | 持平 |
| 5.0 MB | 10.84 ms | 11.54 ms | read 略快（Windows mmap 实现比 Linux 重） |
| 20 MB | 43.69 ms | 46.28 ms | 持平 |

**结论**：mmap 的真正价值**不是速度**（Windows 上甚至略慢 4-6%），而是**省内存 + 大文件不 OOM**。两种路径**命中数完全一致**（行为兼容）。阈值 `MMAP_THRESHOLD_BYTES = 1MB` 写在源码顶部，可按实际场景调整。

## 快速开始

### 1. 标准自检

```bash
cd 23_PDF_Threat_Analyzer
python pdf_threat_analyzer.py --self-check
```

自检会：
1. 探测本地引擎（fitz / pypdf / qpdf / mutool）
2. 手工构造一个含 `/JavaScript` 模拟威胁的测试 PDF
3. 跑一遍 `analyze()`，断言：`is_valid_pdf=True` / 命中 `/JavaScript` / `risk_score > 0` / `risk_level ∈ {MEDIUM, HIGH}`
4. （如有 sanitize 引擎）剥离后重新扫描，断言 `/JavaScript` 已消失

**任何环节失败 → exit 2**。

### 2. 分析指定 PDF

```bash
python pdf_threat_analyzer.py --input "D:\inbox\suspicious.pdf"
python pdf_threat_analyzer.py --input "D:\inbox" --recursive
```

输出到 `output/`：
- `threat_report_<file>.json` —— 单文件报告（含原始关键字命中、可疑 URL、风险分、错误）
- `threat_summary_<时间戳>.txt` —— 批量摘要（按风险等级排序）

### 3. 分析 + 工业级剥离（生成安全 PDF）

```bash
python pdf_threat_analyzer.py --input "D:\inbox" --sanitize
python pdf_threat_analyzer.py --input "D:\inbox" --sanitize --engine fitz --overwrite
```

剥离策略（按优先级）：
1. **fitz (mupdf)**：删除所有注释 / 恶意协议链接 / 嵌入文件 → `garbage=4 + deflate=True` 重写
2. **pypdf 降级**：移除 `OpenAction` / `AA` / `Names.{JavaScript,JS,EmbeddedFiles,Launch}` / `AcroForm`
3. **qpdf 后置校验**（可选）：`qpdf --check` 验证完整性

## 完整参数

```text
--input PATH        输入 PDF 文件或目录（默认: 模块 input/）
--output PATH       输出目录（默认: 模块 output/）
--sanitize          生成 *_sanitized.pdf 安全副本
--engine {auto|fitz|pypdf}
                    sanitize 引擎（默认 auto: fitz 优先 → pypdf 降级）
--recursive         递归遍历子目录（默认）
--no-recursive      不递归
--overwrite         覆盖已存在的输出文件
--self-check        运行标准自检
-v, --verbose       DEBUG 级日志
```

## 风险等级

| 分值 | 等级 | 含义 |
|------|------|------|
| 0 | **LOW** | 未发现威胁特征 |
| 1-4 | **MEDIUM** | 存在可疑对象（外链 / 嵌入文件等） |
| ≥5 | **HIGH** | 明确高危对象（JavaScript / Launch / OpenAction 等） |

权重表见 `pdf_threat_analyzer.py` 内 `RISK_KEYWORDS`：
- `/JavaScript` / `/JS` / `/Launch`：3 分/命中
- `/OpenAction` / `/AA` / `/EmbeddedFiles`：2 分/命中
- `/SubmitForm` / `/URI`：1 分/命中
- 恶意协议 URL（`javascript:` / `data:text/html` / `vbscript:` / `file:`）：2 分/个
- 普通外链：1 分/个（最多 5 分封顶）
- 加密 PDF：+1 分（且无法继续结构化解析）

## 报告示例（节选）

```json
{
  "file_name": "suspicious.pdf",
  "file_size_bytes": 142857,
  "is_valid_pdf": true,
  "is_encrypted": false,
  "raw_keyword_hits": {
    "/JavaScript": 1,
    "/OpenAction": 1
  },
  "extracted_urls": ["https://example.com"],
  "suspicious_urls": [],
  "risk_score": 5,
  "risk_level": "HIGH",
  "errors": []
}
```

## 退出码

遵循 [`docs/logging_convention.md`](../docs/logging_convention.md)：

| Code | 含义 |
|------|------|
| 0 | 成功 |
| 1 | 部分或全部失败 |
| 2 | 自检 / 依赖前置失败（preflight） |
| 130 | 用户中断（Ctrl+C） |

## 日志风格

遵循 [`docs/logging_convention.md`](../docs/logging_convention.md)：

```text
action=analyze_start file=suspicious.pdf size_bytes=142857
action=analyze_done  file=suspicious.pdf risk_level=HIGH score=5
action=sanitize_success file=suspicious_sanitized.pdf engine=fitz(annots=1, links=0, embeds=0)
action=batch_done success=8 failed=0 total=8
```

## 已知限制

1. **加密 PDF**：仅记录为 `is_encrypted=True` 并 +1 分，不做解密后分析
2. **混淆 payload**：基于正则的关键字扫描可被编码/分片绕过；建议二次扫描
3. **表单字段**：未剥离普通表单字段（仅剥离 `AcroForm` 整体，pypdf 路径）
4. **未做 OCR**：扫描的 PDF 必须是文本型（直接读字节）；扫描件请先 OCR
5. **mupdf / qpdf 不一定安装**：未装时按 `fitz → pypdf` 降级；qpdf 缺失仅跳过线性化校验

## 二次扫描建议

模块的"安全 PDF"产出基于**已知威胁字典**的保守剥离。建议在交付前再用：

- **ClamAV**（开源）：`clamscan output/*_sanitized.pdf`
- **Windows Defender**：默认对 PDF 启发式扫描
- **PDFiD**（ Didier Stevens）：`pdfid.py output/*_sanitized.pdf` 验证关键对象已清零
- **VirusTotal API**：上传到多引擎联合扫描

## 环境变量

无强制要求。可选：

- `PYTHONPATH`：标准 Python 包搜索路径
- `PATH`：`qpdf` / `mutool` 需在此路径下（缺失则 graceful skip）
