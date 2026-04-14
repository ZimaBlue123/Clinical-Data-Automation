import re
import sys
from pathlib import Path


BLOCKED_PATTERNS = [
    re.compile(r"SAE_API_TOKEN\s*[:=]\s*['\"][^'\"]{8,}['\"]", re.IGNORECASE),
    re.compile(r"Bearer\s+[A-Za-z0-9._\-]{10,}", re.IGNORECASE),
    re.compile(r"sk-[A-Za-z0-9]{16,}", re.IGNORECASE),
    re.compile(r"(api[_\-]?key|token|secret)\s*[:=]\s*['\"][^'\"]{8,}['\"]", re.IGNORECASE),
]

SKIP_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".webp",
    ".pdf",
    ".zip",
    ".xlsx",
    ".xls",
    ".docx",
    ".doc",
    ".pptx",
}


def should_scan(path: Path) -> bool:
    if not path.exists() or path.is_dir():
        return False
    if path.suffix.lower() in SKIP_EXTENSIONS:
        return False
    if ".git" in path.parts:
        return False
    return True


def scan_file(path: Path) -> list[str]:
    hits: list[str] = []
    try:
        text = path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return hits

    for idx, line in enumerate(text.splitlines(), start=1):
        for pattern in BLOCKED_PATTERNS:
            if pattern.search(line):
                hits.append(f"{path}:{idx}: 检测到疑似密钥/令牌内容")
                break
    return hits


def main() -> int:
    paths = [Path(p) for p in sys.argv[1:]]
    if not paths:
        return 0

    findings: list[str] = []
    for p in paths:
        if should_scan(p):
            findings.extend(scan_file(p))

    if findings:
        print("提交被阻止：发现疑似敏感信息。")
        for f in findings:
            print(f"  - {f}")
        print("\n请改为使用环境变量（例如 SAE_API_TOKEN），不要将密钥写入代码或文档。")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
