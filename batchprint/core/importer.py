"""文件导入 / 扫描：识别 PDF、图片与 Word 文档，递归扫描文件夹，去重。"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Iterable

import fitz  # PyMuPDF

from .settings import PrintJob

log = logging.getLogger(__name__)

PDF_EXTS = {".pdf"}
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".gif"}
WORD_EXTS = {".docx", ".doc"}
SUPPORTED_EXTS = PDF_EXTS | IMAGE_EXTS | WORD_EXTS


def is_supported(path: Path) -> bool:
    # "~$" 开头的是 Word 打开文档时生成的临时锁文件，不是真正的文档
    return path.suffix.lower() in SUPPORTED_EXTS and not path.name.startswith("~$")


def kind_of(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in PDF_EXTS:
        return "pdf"
    if ext in WORD_EXTS:
        return "word"
    return "image"


def _pdf_page_count(path: Path) -> int | None:
    try:
        with fitz.open(path) as doc:
            return doc.page_count
    except Exception as exc:  # noqa: BLE001
        log.warning("读取 PDF 页数失败 %s: %s", path, exc)
        return None


def expand_paths(paths: Iterable[str | Path]) -> list[Path]:
    """把混合的文件/文件夹路径展开成受支持的文件列表（文件夹递归）。"""
    result: list[Path] = []
    seen: set[str] = set()
    for raw in paths:
        p = Path(raw)
        if p.is_dir():
            candidates = sorted(p.rglob("*"))
        else:
            candidates = [p]
        for c in candidates:
            if c.is_file() and is_supported(c):
                key = str(c.resolve()).lower()
                if key not in seen:
                    seen.add(key)
                    result.append(c)
    return result


def build_job(path: Path) -> PrintJob:
    """由单个文件构建打印任务（PDF 会读取页数，属于 I/O，建议放后台线程调用）。

    Word 文档的页数要等转换成 PDF 之后才知道，这里先留空。
    """
    kind = kind_of(path)
    pages = _pdf_page_count(path) if kind == "pdf" else None
    return PrintJob(path=path, kind=kind, pages=pages)


def build_jobs(paths: Iterable[str | Path]) -> list[PrintJob]:
    """由路径列表构建打印任务（读取 PDF 页数）。"""
    return [build_job(path) for path in expand_paths(paths)]
