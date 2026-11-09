"""打印引擎（核心）：用指定 DEVMODE 通过 GDI 打印 PDF / 图片。

管线：
  win32gui.CreateDC("WINSPOOL", printer, devmode)  → 带底层设置的打印机 DC
  → StartDoc / (每页) StartPage → 画图 → EndPage → EndDoc
PDF 用 PyMuPDF 逐页渲染为位图；图片用 Pillow；统一用 PIL.ImageWin 画到 DC。
份数在应用层循环整份文档实现（每份完整输出）。
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Callable, Optional

import fitz  # PyMuPDF
import win32con
import win32gui
import win32ui
from PIL import Image, ImageWin

from . import printers
from .settings import (
    EffectiveSettings,
    FIT_CENTER,
    FIT_FILL,
    FIT_STRETCH,
)

log = logging.getLogger(__name__)

# 渲染 PDF 时的 DPI 上限（兼顾质量与内存）
MAX_RENDER_DPI = 300

# DEVMODE 方向取值
DMORIENT_PORTRAIT = 1
DMORIENT_LANDSCAPE = 2


class PrintCancelled(Exception):
    """用户中止打印。"""


CancelCheck = Optional[Callable[[], bool]]


def parse_page_range(spec: Optional[str], total: int) -> list[int]:
    """把 "1-3,5" 解析成 0 基页码列表；spec 空则返回全部页。"""
    if not spec or not spec.strip():
        return list(range(total))
    pages: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, _, b = part.partition("-")
            try:
                start = int(a)
                end = int(b)
            except ValueError:
                continue
            for n in range(start, end + 1):
                if 1 <= n <= total:
                    pages.append(n - 1)
        else:
            try:
                n = int(part)
            except ValueError:
                continue
            if 1 <= n <= total:
                pages.append(n - 1)
    # 去重保序
    seen: set[int] = set()
    result: list[int] = []
    for n in pages:
        if n not in seen:
            seen.add(n)
            result.append(n)
    return result or list(range(total))


def _compute_rect(img_w: int, img_h: int, area_w: int, area_h: int,
                  fit: str) -> tuple[int, int, int, int]:
    """计算图片在可打印区域内的目标矩形 (x1, y1, x2, y2)。"""
    if fit == FIT_STRETCH:
        return (0, 0, area_w, area_h)

    if fit == FIT_FILL:
        # 等比放大到铺满，居中（超出部分被裁切）
        scale = max(area_w / img_w, area_h / img_h)
    else:  # FIT_CENTER 默认：等比缩放，完整显示
        scale = min(area_w / img_w, area_h / img_h)

    w = max(1, int(round(img_w * scale)))
    h = max(1, int(round(img_h * scale)))
    x = (area_w - w) // 2
    y = (area_h - h) // 2
    return (x, y, x + w, y + h)


def _draw_image_on_dc(dc: "win32ui.PyCDC", img: Image.Image, fit: str) -> None:
    """把一张 PIL 图片按 fit 模式画到打印机 DC 的一页上。"""
    if img.mode != "RGB":
        img = img.convert("RGB")

    area_w = dc.GetDeviceCaps(win32con.HORZRES)   # 可打印区宽（像素）
    area_h = dc.GetDeviceCaps(win32con.VERTRES)   # 可打印区高（像素）

    rect = _compute_rect(img.width, img.height, area_w, area_h, fit)
    dib = ImageWin.Dib(img)
    dib.draw(dc.GetHandleOutput(), rect)


def _iter_pdf_images(path: Path, page_indices: list[int], render_dpi: int):
    """逐页把 PDF 渲染成 PIL 图片（生成器，渲染完即释放）。"""
    zoom = render_dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)
    with fitz.open(path) as doc:
        for idx in page_indices:
            if idx < 0 or idx >= doc.page_count:
                continue
            page = doc.load_page(idx)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            yield img
            del pix
            del img


def _create_printer_dc(printer_name: str, devmode: Optional[Any]) -> "win32ui.PyCDC":
    """用给定 DEVMODE 创建打印机 DC。"""
    hdc = win32gui.CreateDC("WINSPOOL", printer_name, devmode)
    return win32ui.CreateDCFromHandle(hdc)


def _check_cancel(cancel: CancelCheck) -> None:
    if cancel and cancel():
        raise PrintCancelled()


def _orientation_for(w: float, h: float) -> int:
    return DMORIENT_LANDSCAPE if w > h else DMORIENT_PORTRAIT


def _pdf_page_indices(path: Path, settings: EffectiveSettings) -> list[int]:
    with fitz.open(path) as doc:
        total = doc.page_count
    return parse_page_range(settings.page_range, total)


def _job_orientations(path: Path, kind: str,
                      settings: EffectiveSettings) -> list[int]:
    """按内容宽高比算出每页/每图应有的方向（不渲染，仅读尺寸）。"""
    if kind == "pdf":
        indices = _pdf_page_indices(path, settings)
        out: list[int] = []
        with fitz.open(path) as doc:
            for idx in indices:
                r = doc.load_page(idx).rect
                out.append(_orientation_for(r.width, r.height))
        return out
    with Image.open(path) as img:
        w, h = img.size
    return [_orientation_for(w, h)]


def _iter_job_images(path: Path, kind: str, settings: EffectiveSettings,
                     render_dpi: int):
    """按页顺序产出 PIL 图片（与 _job_orientations 顺序一致）。"""
    if kind == "pdf":
        indices = _pdf_page_indices(path, settings)
        yield from _iter_pdf_images(path, indices, render_dpi)
    else:
        with Image.open(path) as img:
            img.load()
            yield img


def _mutable_devmode(settings: EffectiveSettings):
    """返回一个可安全修改方向的独立 DEVMODE；无法获得返回 None。

    若已有 devmode 则克隆（不改到存储对象）；否则取打印机默认（本身即独立对象）。
    """
    if settings.devmode is not None:
        return printers.clone_devmode(settings.printer_name, settings.devmode)
    return printers.get_default_devmode(settings.printer_name)


def print_job(settings: EffectiveSettings, path: Path, kind: str,
              cancel: CancelCheck = None, output_file: str | None = None,
              doc_name: str | None = None) -> None:
    """用生效设置打印单个文件；抛异常表示失败，PrintCancelled 表示被中止。

    output_file: 若指定，则把输出重定向到该文件（用于"打印到文件"驱动或测试），
                 而不是正常送打印机。通常留空。
    doc_name:    打印队列里显示的文档名。Word 文档实际打印的是缓存里的 PDF，
                 用这个参数显示原文件名，而不是缓存文件名。
    """
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {path}")
    if not settings.printer_name:
        raise ValueError("未选择打印机")

    copies = max(1, int(settings.copies or 1))

    # 自动方向：算出每页方向，并准备一个可逐页修改的独立 DEVMODE
    auto = settings.auto_orient
    orientations = _job_orientations(path, kind, settings) if auto else []
    work_devmode = _mutable_devmode(settings) if (auto and orientations) else None
    dc_devmode = settings.devmode
    if work_devmode is not None:
        try:
            work_devmode.Orientation = orientations[0]
            dc_devmode = work_devmode          # 首页方向用于建 DC
        except Exception:                       # noqa: BLE001
            work_devmode = None
    auto = work_devmode is not None

    log.info("打印 %s（%s，%d 份，自动方向=%s）到 %s", doc_name or path.name, kind,
             copies, auto, settings.printer_name)

    dc = _create_printer_dc(settings.printer_name, dc_devmode)
    try:
        render_dpi = min(dc.GetDeviceCaps(win32con.LOGPIXELSX) or 300, MAX_RENDER_DPI)
        current_orient = orientations[0] if auto else None

        title = doc_name or str(path)
        for _copy in range(copies):
            _check_cancel(cancel)
            if output_file:
                dc.StartDoc(title, output_file)
            else:
                dc.StartDoc(title)
            try:
                for page_no, img in enumerate(
                    _iter_job_images(path, kind, settings, render_dpi)
                ):
                    _check_cancel(cancel)
                    if auto and page_no < len(orientations):
                        want = orientations[page_no]
                        if want != current_orient:
                            work_devmode.Orientation = want
                            win32gui.ResetDC(dc.GetSafeHdc(), work_devmode)
                            current_orient = want
                    dc.StartPage()
                    _draw_image_on_dc(dc, img, settings.image_fit)
                    dc.EndPage()
                dc.EndDoc()
            except Exception:
                dc.AbortDoc()
                raise
    finally:
        dc.DeleteDC()
