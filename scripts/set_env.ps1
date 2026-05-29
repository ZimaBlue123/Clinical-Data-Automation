$ErrorActionPreference = "Stop"

# 用法（在仓库根目录）：
#   . .\scripts\set_env.ps1
# 然后进入 32_SAE_Extractor 运行 cli.py 等。

$RepoRoot = Split-Path $PSScriptRoot -Parent

if ([string]::IsNullOrWhiteSpace($env:SAE_API_TOKEN)) {
  Write-Host "请先设置 SAE_API_TOKEN（必填）。例如：" -ForegroundColor Yellow
  Write-Host "  `$env:SAE_API_TOKEN='你的Token'" -ForegroundColor Yellow
}

# 本地网关（常见为 SSH 隧道暴露的地址，可按需修改）
if ([string]::IsNullOrWhiteSpace($env:SAE_API_BASE_URL)) {
  $env:SAE_API_BASE_URL = "http://127.0.0.1:10984"
}

# 默认输出目录：第 31 模块 output（可通过 SAE_OUTPUT_DIR 覆盖）
if ([string]::IsNullOrWhiteSpace($env:SAE_OUTPUT_DIR)) {
  $env:SAE_OUTPUT_DIR = Join-Path $RepoRoot "32_SAE_Extractor\output"
}

Write-Host "已设置：SAE_API_BASE_URL=$env:SAE_API_BASE_URL" -ForegroundColor Green
Write-Host "已设置：SAE_OUTPUT_DIR=$env:SAE_OUTPUT_DIR" -ForegroundColor Green
