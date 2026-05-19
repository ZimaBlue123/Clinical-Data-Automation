# 28_DNS_Leak_Detector

DNS 泄漏诊断模块（支持 `tun` / `socks` 两种模式）。

## 功能

- 检测公网出口地域与上游 DNS 解析地域是否偏离
- 输出风险结论：正常 / 潜在偏移 / 高危泄漏
- 可将诊断结果导出为 JSON 报告到 `output/`

## 快速开始

```bash
cd 28_DNS_Leak_Detector
python dns_leak_detector.py --mode tun
```

## 常用参数

```bash
# SOCKS 模式（默认 10808）
python dns_leak_detector.py --mode socks --socks-port 10808

# 调整请求超时
python dns_leak_detector.py --timeout 10

# 保存 JSON 报告
python dns_leak_detector.py --save-json
```

## 输出说明

- 诊断日志：终端实时输出
- JSON 报告（可选）：`output/dns_diagnostic_<mode>_<timestamp>.json`

