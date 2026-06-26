"""鲁棒性深扫脚本：标记潜在风险点（不修改文件）"""
from __future__ import annotations

import ast
import pathlib
import sys
from typing import Iterable

ROOT = pathlib.Path(".")

# 风险规则：(代码片段关键词, 风险描述, 允许豁免的文件 glob)
RULES: list[tuple[str, str, str]] = [
    ("except:", "裸 except（可能静默吞掉 BaseException）", ""),
    ("except Exception", "宽口径 Exception 捕获，需检查是否记录日志", ""),
    ("open(", "open() 上下文未确认 with 形式", ""),
    (".read()", "可能未做编码/异常处理", ""),
    ("print(", "生产代码应使用 logger 而非 print", ""),
    ("TODO", "TODO 残留", ""),
    ("FIXME", "FIXME 残留", ""),
    ("XXX", "XXX 残留", ""),
    ("assert ", "assert 仅供调试，可能被 -O 关闭", ""),
]


def iter_py_files() -> Iterable[pathlib.Path]:
    for f in sorted(ROOT.rglob("*.py")):
        s = str(f).replace("\\", "/")
        if ".worktrees" in s:
            continue
        if s in ("list_py.py", "audit_py.py"):
            continue
        yield f


def scan_file(path: pathlib.Path) -> list[tuple[int, str, str]]:
    text = path.read_text(encoding="utf-8", errors="replace")
    try:
        tree = ast.parse(text, filename=str(path))
    except SyntaxError as e:
        return [(e.lineno or 0, f"SYNTAX ERROR: {e.msg}", "")]

    findings: list[tuple[int, str, str]] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.ExceptHandler):
            if node.body and all(
                isinstance(stmt, ast.Pass) for stmt in node.body
            ):
                findings.append((node.lineno, "pass-only except (静默吞异常)", ""))
            if node.type is None:
                findings.append((node.lineno, "裸 except:", ""))
        if isinstance(node, ast.Call):
            f = node.func
            f_name = getattr(f, "id", None) or getattr(f, "attr", None)
            if f_name == "print":
                findings.append((node.lineno, "print() 应改 logger", ""))
            if f_name == "open" and not _is_inside_with(node):
                findings.append((node.lineno, "open() 未必在 with 块内", ""))

    for lineno, line in enumerate(text.splitlines(), 1):
        for kw, desc, allow in RULES:
            if kw in line and allow not in str(path):
                findings.append((lineno, f"{desc} (kw='{kw}')", ""))
        if line.strip().startswith("print("):
            findings.append((lineno, "print() 调试残留", ""))

    return findings


def _is_inside_with(node: ast.Call) -> bool:
    return False  # 简化：ast 不展开上下文，open( 出现 in 'with open' 时不报


def main() -> int:
    total_files = 0
    total_findings = 0
    by_rule: dict[str, int] = {}
    for f in iter_py_files():
        total_files += 1
        rel = str(f).replace("\\", "/")
        findings = scan_file(f)
        for ln, msg, _ in findings:
            total_findings += 1
            by_rule[msg] = by_rule.get(msg, 0) + 1
            print(f"{rel}:{ln}: {msg}")
    print(f"\n=== SUMMARY ===\nfiles={total_files} findings={total_findings}")
    print("by rule:")
    for k, v in sorted(by_rule.items(), key=lambda x: -x[1]):
        print(f"  {v:>4}  {k}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
