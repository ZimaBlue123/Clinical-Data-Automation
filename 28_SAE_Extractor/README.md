# 28_SAE_Extractor

从临床文档（PDF / TXT / DOCX / Excel）中抽取严重不良事件（SAE）结构化字段，经 OpenAI 兼容 API 解析后导出 Excel。

## 输出字段

`sae_term`、`onset_date`、`resolution_date`、`severity_grade`、`causality`、`action_taken`、`outcome`，批处理时附加 `source_file`。

## 环境

- Python 3.8+（与仓库一致；API 与 OCR 链路建议 3.9+）
- 可选：本机 Tesseract、Poppler（PDF OCR 回退）
- 必须：可访问的 Chat Completions 网关与 `SAE_API_TOKEN`

## 配置（PowerShell 示例）

```powershell
$env:SAE_API_TOKEN="你的Token"
$env:SAE_API_BASE_URL="http://127.0.0.1:10984"
$env:SAE_MODEL_ID="你的模型ID"
$env:SAE_INPUT_DIR="E:\path\to\input"
$env:SAE_OUTPUT_DIR="E:\path\to\output"
$env:TESSERACT_CMD="C:\Program Files\Tesseract-OCR\tesseract.exe"
$env:POPPLER_PATH="D:\poppler\Library\bin"
```

## 用法

默认目录：`28_SAE_Extractor/input/` → `28_SAE_Extractor/output/`。

```bash
cd 28_SAE_Extractor
python cli.py self-check
python cli.py batch
python cli.py batch --fast-docx
python cli.py pdf-batch
python cli.py batch --input-dir ".\input" --output ".\output\listing.xlsx"
```

单文件内置样例（需已配置 Token）：

```bash
python sae_extractor.py
```

依赖自检：

```bash
python sae_env_check.py
```

## 安全

勿将 Token 写入代码或提交仓库；处理前建议对受试者信息脱敏。

仓库根目录提供 **`scripts/check_secrets.py`**（由 **pre-commit** 在提交前调用，配置见项目根 `.pre-commit-config.yaml`）。可执行 `pip install pre-commit` → `pre-commit install`，或手动 `pre-commit run --all-files`。

远程 API 需隧道时，可使用根目录 **`scripts/start_tunnel.ps1`**（须设置 `SAE_TUNNEL_SSH_HOST` 等）；本模块环境变量示例：**`scripts/set_env.ps1`**。
