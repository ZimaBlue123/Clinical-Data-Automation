from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
import os
import platform
import sys
import time

import fitz


# 重要：translators 在网络/地区检测失败时会报错要求设置该变量。
os.environ.setdefault("translators_default_region", "CN")
import translators as ts  # type: ignore


BASE = Path(__file__).resolve().parent
INPUT_PDF = BASE / "input" / "inclusion-exclusion-criteria-important-medical-events-ime-list_en.pdf"
OUT_DIR = BASE / "output"
OUTPUT_PDF = OUT_DIR / "inclusion-exclusion-criteria-important-medical-events-ime-list_en_en2zh_FINAL_REPLACE.pdf"


def _resolve_cjk_fontfile() -> str:
    env = os.environ.get("PDF_CJK_FONT", "").strip()
    if env and Path(env).is_file():
        return env
    if platform.system().lower() == "windows":
        windir = Path(os.environ.get("WINDIR", r"C:\Windows"))
        for name in ("msyh.ttc", "msyhbd.ttc", "simsun.ttc", "simhei.ttf"):
            p = windir / "Fonts" / name
            if p.is_file():
                return str(p)
    raise FileNotFoundError("未找到可用中文字体：请设置环境变量 PDF_CJK_FONT 指向 .ttf/.ttc")


def _is_good_translation(src: str, tr: str) -> bool:
    if not tr:
        return False
    tr = tr.strip()
    if not tr or tr == src.strip():
        return False
    # 优先出现中文字符
    has_cjk = any("\u4e00" <= ch <= "\u9fff" for ch in tr)
    if has_cjk:
        return True
    # 否则至少确保不是纯空白/原样
    return any(ord(ch) > 127 for ch in tr)


@dataclass(frozen=True)
class SpanItem:
    rect: fitz.Rect
    fontsize: float
    src: str


def _translate_unique(
    texts: Iterable[str],
    cache: dict[str, str | None],
    sleep_s: float,
    max_attempts: int = 4,
    engine: str = "bing",
) -> None:
    for s in texts:
        if s is None:
            continue
        key = str(s).strip()
        if not key or key in cache:
            continue

        tr: str | None = None
        for attempt in range(max_attempts):
            try:
                cand = ts.translate_text(
                    key[:4800],
                    translator=engine,
                    from_language="en",
                    to_language="zh",
                )
                cand = str(cand).strip() if cand is not None else ""
                if _is_good_translation(key, cand):
                    tr = cand
                    break
            except Exception:
                pass
            time.sleep(sleep_s * (attempt + 1))
        cache[key] = tr
        time.sleep(sleep_s)


def translate_pdf_overlay() -> Path:
    fontfile = _resolve_cjk_fontfile()
    sleep_s = float(os.environ.get("TRANSLATE_SLEEP", "0.45"))

    cache: dict[str, str | None] = {}

    doc = fitz.open(str(INPUT_PDF))
    for pno, page in enumerate(doc, start=1):
        info = page.get_text("dict")
        items: list[SpanItem] = []
        unique_texts: list[str] = []

        for block in info.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    txt = (span.get("text") or "").strip()
                    if not txt:
                        continue
                    bbox = fitz.Rect(span["bbox"])
                    pad = 0.25
                    rect = fitz.Rect(bbox.x0 - pad, bbox.y0 - pad, bbox.x1 + pad, bbox.y1 + pad)
                    rect_h = max(1e-6, rect.y1 - rect.y0)
                    fontsize_raw = float(span.get("size", 10) or 10)

                    # 关键：span bbox 高度往往很小，小矩形下 insert_textbox 容易直接不渲染。
                    # 这里把字号按 bbox 高度压缩，并确保 bbox 高度至少是字号的 2 倍。
                    fontsize = max(2.5, min(fontsize_raw, rect_h * 0.6))
                    min_rect_h = fontsize * 2.0
                    if rect_h < min_rect_h:
                        extra = (min_rect_h - rect_h) / 2.0
                        rect = fitz.Rect(rect.x0, rect.y0 - extra, rect.x1, rect.y1 + extra)
                    items.append(SpanItem(rect=rect, fontsize=fontsize, src=txt))
                    unique_texts.append(txt)

        if not items:
            continue

        _translate_unique(unique_texts, cache, sleep_s=sleep_s)

        if pno == 1:
            for it in items[:6]:
                tr = cache.get(it.src)
                print(f"[debug] src='{it.src[:45]}' => tr='{(tr or '')[:45]}'", flush=True)

        for it in items:
            tr = cache.get(it.src)
            if not tr:
                continue
            if tr.strip() == it.src.strip():
                continue

            rect0 = it.rect
            # 替换模式：先把原文区域覆盖成白底（overlay=True），再把翻译文字写到同一 bbox。
            # 通过在每次重试前都覆盖该 bbox，避免“重影/残留原文”。
            for scale in (1.0, 0.85, 0.72, 0.60, 0.50):
                fs = max(2.5, it.fontsize * scale)
                pad_retry = max(0.05, 0.08 * (1.0 + (1.0 - scale)))
                rect = fitz.Rect(
                    rect0.x0 - pad_retry,
                    rect0.y0 - pad_retry,
                    rect0.x1 + pad_retry,
                    rect0.y1 + pad_retry,
                )

                # 先覆盖，确保原文不再可见（避免你截图里的“中英重影”）
                page.draw_rect(
                    rect,
                    color=(1, 1, 1),
                    fill=(1, 1, 1),
                    width=0,
                    overlay=True,
                    stroke_opacity=0,
                    fill_opacity=1,
                )

                rv = page.insert_textbox(
                    rect,
                    tr,
                    fontsize=fs,
                    fontfile=fontfile,
                    fontname="f0",
                    color=(0, 0, 0),
                    align=0,
                    fill_opacity=1,
                    overlay=True,
                )
                if rv >= 0:
                    break

        print(f"[page] {pno}/{doc.page_count} segments={len(items)}", flush=True)

    OUTPUT_PDF.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(OUTPUT_PDF))
    doc.close()
    return OUTPUT_PDF


def self_check_text(pdf_path: Path) -> tuple[int, int]:
    d = fitz.open(str(pdf_path))
    full = "".join(page.get_text("text") for page in d)
    d.close()
    cjk = sum(1 for ch in full if "\u4e00" <= ch <= "\u9fff")
    return cjk, len(full)


def self_check_pixmap_diff(src_pdf: Path, out_pdf: Path, max_pages: int = 2) -> None:
    import numpy as np

    sa = fitz.open(str(src_pdf))
    sb = fitz.open(str(out_pdf))
    n = min(max_pages, sa.page_count, sb.page_count)
    for i in range(n):
        x = sa.load_page(i).get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        y = sb.load_page(i).get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        a = np.frombuffer(x.samples, dtype=np.uint8).reshape(x.height, x.width, x.n)
        b = np.frombuffer(y.samples, dtype=np.uint8).reshape(y.height, y.width, y.n)
        mean_abs = float(np.mean(np.abs(a.astype(np.int16) - b.astype(np.int16))))
        changed_ratio = float(np.mean(np.any(np.abs(a.astype(np.int16) - b.astype(np.int16)) > 10, axis=2)))
        print(f"[diff] page {i+1}: mean_abs={mean_abs:.4f}, changed_ratio={changed_ratio:.6f}", flush=True)
    sa.close()
    sb.close()


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass

    out = translate_pdf_overlay()
    cjk, total = self_check_text(out)
    print(f"OUTPUT={out}")
    print(f"CJK_COUNT={cjk}, TEXT_LEN={total}")
    self_check_pixmap_diff(INPUT_PDF, out, max_pages=2)

