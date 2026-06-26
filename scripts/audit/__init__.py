"""scripts/audit/ — 仓库级代码与配置审计工具。

本子包提供可复用的静态分析脚本：

- audit_py: 鲁棒性深扫（标记潜在风险点，不修改文件）
- audit_deps: 依赖一致性审计（解析 requirements*.txt vs 代码 import）

调用方式（从仓库根）：
    python -m scripts.audit.audit_py
    python -m scripts.audit.audit_deps
"""
