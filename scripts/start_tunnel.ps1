$ErrorActionPreference = "Stop"

# SSH 本地端口转发（默认：本机端口 -> 远端 127.0.0.1:10984）
# 用法（在仓库根目录）：
#   powershell -ExecutionPolicy Bypass -File .\scripts\start_tunnel.ps1
#
# 必填环境变量：
#   SAE_TUNNEL_SSH_HOST   远端主机（示例：your.server.example）
# 可选：
#   SAE_TUNNEL_SSH_USER       默认 root
#   SAE_TUNNEL_LOCAL_PORT    默认 10984
#   SAE_TUNNEL_REMOTE_BIND    默认 127.0.0.1:10984

$LocalPort = $env:SAE_TUNNEL_LOCAL_PORT
if ([string]::IsNullOrWhiteSpace($LocalPort)) { $LocalPort = "10984" }

$RemoteHost = $env:SAE_TUNNEL_SSH_HOST
if ([string]::IsNullOrWhiteSpace($RemoteHost)) {
  Write-Host "请先设置 SAE_TUNNEL_SSH_HOST（远端 SSH 主机名或 IP）。" -ForegroundColor Red
  Write-Host "  例如：`$env:SAE_TUNNEL_SSH_HOST='user@vps.example'" -ForegroundColor Yellow
  exit 1
}

$RemoteUser = $env:SAE_TUNNEL_SSH_USER
if ([string]::IsNullOrWhiteSpace($RemoteUser)) { $RemoteUser = "root" }

$RemoteBind = $env:SAE_TUNNEL_REMOTE_BIND
if ([string]::IsNullOrWhiteSpace($RemoteBind)) { $RemoteBind = "127.0.0.1:10984" }

Write-Host "启动 SSH 隧道: 127.0.0.1:$LocalPort  ->  $RemoteUser@$RemoteHost ($RemoteBind)" -ForegroundColor Cyan
Write-Host "提示：保持此窗口不关闭；要退出请按 Ctrl+C。" -ForegroundColor Yellow

ssh -N -L "$LocalPort`:$RemoteBind" "$RemoteUser@$RemoteHost"
