"""依赖审计。"""
from __future__ import annotations
import ast
import pathlib
import re
from collections import defaultdict

def parse_req(p: pathlib.Path) -> set[str]:
    if not p.exists():
        return set()
    out: set[str] = set()
    for line in p.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        m = re.match(r"^([A-Za-z0-9_.\-]+)", line)
        if m:
            out.add(m.group(1).lower())
    return out

req_full = parse_req(pathlib.Path("requirements.txt"))
req_lite = parse_req(pathlib.Path("requirements.lite.txt"))
req_ci = parse_req(pathlib.Path("requirements-ci.txt"))
req_all = req_full | req_lite | req_ci

imports: set[str] = set()
import_to_files: dict[str, list[str]] = defaultdict(list)
SKIP = {".worktrees"}
EXCLUDE = {
    "_audit_deps.py", "_fix_sae.py", "_fix_excel.py", "_verify.py",
    "_delete_caches.py", "_scan_junk.py", "list_py.py", "audit_py.py",
}
for p in pathlib.Path(".").rglob("*.py"):
    s = str(p).replace("\\", "/")
    if any(seg in s for seg in SKIP):
        continue
    if p.name in EXCLUDE:
        continue
    try:
        tree = ast.parse(p.read_text(encoding="utf-8", errors="ignore"))
    except SyntaxError:
        continue
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for n in node.names:
                pkg = n.name.split(".")[0]
                imports.add(pkg.lower())
                import_to_files[pkg.lower()].append(s)
        elif isinstance(node, ast.ImportFrom):
            if node.module and node.level == 0:
                pkg = node.module.split(".")[0]
                imports.add(pkg.lower())
                import_to_files[pkg.lower()].append(s)

PACKAGE_MAP = {
    "sklearn": "scikit-learn",
    "yaml": "pyyaml",
    "pil": "pillow",
    "fitz": "pymupdf",
    "cv2": "opencv-python",
    "pptx": "python-pptx",
    "docx": "python-docx",
    "win32com": "pywin32",
}
normalized_imports = {PACKAGE_MAP.get(p, p).lower() for p in imports}

declared_but_not_imported = sorted(req_all - normalized_imports)
imported_but_not_declared = sorted(normalized_imports - req_all)

stdlib = {
    "argparse", "ast", "asyncio", "base64", "collections", "concurrent", "configparser",
    "contextlib", "csv", "ctypes", "dataclasses", "datetime", "decimal", "difflib",
    "enum", "errno", "fnmatch", "functools", "gc", "getopt", "getpass", "glob",
    "gzip", "hashlib", "heapq", "html", "http", "importlib", "inspect", "io",
    "ipaddress", "itertools", "json", "logging", "math", "mimetypes", "multiprocessing",
    "numbers", "operator", "os", "pathlib", "pickle", "platform", "posixpath",
    "pprint", "queue", "random", "re", "shlex", "shutil", "signal", "smtplib",
    "socket", "sqlite3", "ssl", "stat", "statistics", "string", "struct",
    "subprocess", "sys", "tempfile", "textwrap", "threading", "time", "timeit",
    "tokenize", "traceback", "types", "typing", "unicodedata", "unittest",
    "urllib", "uuid", "venv", "warnings", "weakref", "xml", "zipfile", "zlib",
    "builtins", "__future__", "typing_extensions", "tomllib", "email",
    "base64", "copy", "pdb", "profile", "pstats",
    "stringprep", "codecs", "calendar", "bisect", "array",
    "fileinput", "cgi", "wsgiref", "xmlrpc", "pipes",
}

truly_undeclared = [p for p in imported_but_not_declared if p not in stdlib]

log: list[str] = []
log.append("=" * 70)
log.append("依赖审计结果")
log.append("=" * 70)
log.append(f"requirements.txt 声明: {len(req_full)} 个")
log.append(f"requirements.lite.txt 声明: {len(req_lite)} 个")
log.append(f"requirements-ci.txt 声明: {len(req_ci)} 个")
log.append(f"全仓实际 import: {len(normalized_imports)} 个")
log.append("")
log.append("--- A. 声明了但未在代码中发现 import ---")
for p in declared_but_not_imported:
    log.append(f"  {p}")
log.append("")
log.append("--- B. 实际 import 但未在 requirements 中声明（剔除标准库）---")
for p in truly_undeclared:
    files = import_to_files.get(p, [])[:5]
    log.append(f"  {p}  -- files: {files}")
log.append("")
log.append(f"--- C. 重要映射检查 ---")
log.append(f"  yaml  -> pyyaml: {('pyyaml' in normalized_imports)}")
log.append(f"  PIL   -> pillow: {('pillow' in normalized_imports)}")
log.append(f"  fitz  -> pymupdf: {('pymupdf' in normalized_imports)}")

pathlib.Path("deps_audit.txt").write_text("\n".join(log), encoding="utf-8")
print("done")
