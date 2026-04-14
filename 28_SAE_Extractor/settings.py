import os
import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Dict, Any, Tuple


@dataclass(frozen=True)
class Settings:
    project_root: Path
    input_dir: Path
    output_dir: Path
    api_base_url: str
    api_token: Optional[str]
    model_id: str
    tesseract_cmd: Optional[str]
    poppler_path: Optional[str]


def get_project_root() -> Path:
    return Path(__file__).resolve().parent


def _coerce_path(value: Optional[str], default: Path, *, root: Path) -> Path:
    """
    将环境变量/参数解析为稳定的绝对 Path：
    - 支持 %VAR% 展开（由 OS 负责注入到进程环境后，这里做 expandvars）
    - 相对路径按模块根目录 root 解析
    """
    if not value:
        return default
    expanded = os.path.expandvars(value).strip()
    p = Path(expanded).expanduser()
    if not p.is_absolute():
        p = root / p
    return p.resolve()


def load_settings() -> Settings:
    root = get_project_root()

    input_dir = _coerce_path(os.environ.get("SAE_INPUT_DIR"), root / "input", root=root)
    output_dir = _coerce_path(os.environ.get("SAE_OUTPUT_DIR"), root / "output", root=root)

    api_base_url = os.environ.get("SAE_API_BASE_URL", "http://127.0.0.1:10984")
    api_token = os.environ.get("SAE_API_TOKEN")
    model_id = os.environ.get("SAE_MODEL_ID", "doubao-seed-2-0-pro-260215")

    tesseract_cmd = os.environ.get("TESSERACT_CMD") or os.environ.get("SAE_TESSERACT_CMD")
    poppler_path = os.environ.get("POPPLER_PATH") or os.environ.get("SAE_POPPLER_PATH")

    return Settings(
        project_root=root,
        input_dir=input_dir,
        output_dir=output_dir,
        api_base_url=api_base_url,
        api_token=api_token,
        model_id=model_id,
        tesseract_cmd=tesseract_cmd,
        poppler_path=poppler_path,
    )


def ensure_dirs(settings: Settings) -> None:
    settings.input_dir.mkdir(parents=True, exist_ok=True)
    settings.output_dir.mkdir(parents=True, exist_ok=True)


def resolve_default_output(settings: Settings, filename: str) -> Path:
    return settings.output_dir / filename


def _run_version_cmd(args: list[str]) -> Tuple[bool, str]:
    try:
        p = subprocess.run(args, capture_output=True, text=True, timeout=10)
        first_line = (p.stdout or p.stderr or "").splitlines()[:1]
        return (p.returncode == 0, first_line[0] if first_line else "")
    except FileNotFoundError:
        return (False, "")
    except Exception:
        return (False, "")


def check_environment(settings: Settings) -> Dict[str, Any]:
    result: Dict[str, Any] = {
        "python": {"ok": True},
        "tesseract": {"ok": False, "source": None, "detail": ""},
        "poppler": {"ok": False, "source": None, "detail": ""},
        "api": {"base_url": settings.api_base_url, "token_set": bool(settings.api_token), "model_id": settings.model_id},
        "paths": {
            "project_root": str(settings.project_root),
            "input_dir": str(settings.input_dir),
            "output_dir": str(settings.output_dir),
        },
    }

    if settings.tesseract_cmd:
        if Path(settings.tesseract_cmd).exists():
            ok, detail = _run_version_cmd([settings.tesseract_cmd, "--version"])
            result["tesseract"] = {"ok": ok, "source": "TESSERACT_CMD", "detail": detail}
        else:
            result["tesseract"] = {"ok": False, "source": "TESSERACT_CMD", "detail": f"路径不存在: {settings.tesseract_cmd}"}
    else:
        which = shutil.which("tesseract")
        if which:
            ok, detail = _run_version_cmd(["tesseract", "--version"])
            result["tesseract"] = {"ok": ok, "source": "PATH", "detail": detail}

    if settings.poppler_path:
        candidate = Path(settings.poppler_path) / ("pdftoppm.exe" if os.name == "nt" else "pdftoppm")
        if candidate.exists():
            ok, detail = _run_version_cmd([str(candidate), "-v"])
            result["poppler"] = {"ok": ok, "source": "POPPLER_PATH", "detail": detail}
        else:
            result["poppler"] = {"ok": False, "source": "POPPLER_PATH", "detail": f"未找到: {candidate}"}
    else:
        which = shutil.which("pdftoppm")
        if which:
            ok, detail = _run_version_cmd(["pdftoppm", "-v"])
            result["poppler"] = {"ok": ok, "source": "PATH", "detail": detail}

    return result
