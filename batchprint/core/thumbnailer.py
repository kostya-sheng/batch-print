"""缩略图生成：PDF 取首页，图片直接缩放；在后台线程池执行，避免卡界面。

要点（踩过的坑）：
- QPixmap 只能在 GUI 线程使用。工作线程里只产出 QImage（线程安全），
  QPixmap 一律在主线程转换，否则会与主线程争锁导致界面"未响应"。
- 用专用线程池并限制并发，避免大批缩略图渲染占满所有 CPU 核心把 UI 线程饿死。
"""
from __future__ import annotations

import logging
import os
from pathlib import Path

import fitz  # PyMuPDF
from PySide6.QtCore import QObject, QRunnable, Qt, QThreadPool, Signal, Slot
from PySide6.QtGui import QImage, QImageReader, QPixmap

log = logging.getLogger(__name__)


def _render_pdf_first_page(path: Path, max_px: int) -> QImage | None:
    try:
        with fitz.open(path) as doc:
            if doc.page_count == 0:
                return None
            page = doc.load_page(0)
            rect = page.rect
            scale = max_px / max(rect.width, rect.height, 1)
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            img = QImage(pix.samples, pix.width, pix.height,
                         pix.stride, QImage.Format_RGB888)
            return img.copy()  # 复制，脱离 pix 内存
    except Exception as exc:  # noqa: BLE001
        log.warning("生成 PDF 缩略图失败 %s: %s", path, exc)
        return None


def _render_image(path: Path, max_px: int) -> QImage | None:
    try:
        reader = QImageReader(str(path))
        size = reader.size()
        if size.isValid() and (size.width() > max_px or size.height() > max_px):
            # 关键：让解码器直接按缩小后的尺寸解码，
            # 不要把整张大图（如 4000x3000）解进内存再缩——那样又慢又占内存
            reader.setScaledSize(size.scaled(max_px, max_px, Qt.KeepAspectRatio))
        img = reader.read()
        if img.isNull():
            return None
        return img
    except Exception as exc:  # noqa: BLE001
        log.warning("生成图片缩略图失败 %s: %s", path, exc)
        return None


class _Signals(QObject):
    # (key, 缩略图)。用稳定的 key(文件路径) 而非行号——行会因删除而移位，
    # 用行号会把缩略图贴到错误的行上。只传 QImage，QPixmap 不能跨线程。
    done = Signal(str, QImage)


class _ThumbTask(QRunnable):
    def __init__(self, key: str, path: Path, kind: str, max_px: int,
                 signals: _Signals):
        super().__init__()
        self.key = key
        self.path = path
        self.kind = kind
        self.max_px = max_px
        self.signals = signals

    def run(self) -> None:
        if self.kind == "pdf":
            img = _render_pdf_first_page(self.path, self.max_px)
        else:
            img = _render_image(self.path, self.max_px)
        # 失败也要发（空 QImage），否则调用方的进度计数会一直挂着
        self.signals.done.emit(self.key, img if img is not None else QImage())


class Thumbnailer(QObject):
    """异步缩略图服务；ready 信号在主线程发出，可安全更新 UI。"""

    ready = Signal(str, QPixmap)

    def __init__(self, max_px: int = 96, parent=None):
        super().__init__(parent)
        # 专用线程池并限制并发：留出 CPU 给 UI 线程，避免批量导入时界面卡死
        self._pool = QThreadPool()
        self._pool.setMaxThreadCount(max(1, min(4, (os.cpu_count() or 2) // 2)))
        self._max_px = max_px
        self._signals = _Signals()
        # _signals 属于主线程 -> 跨线程 emit 走队列连接，_on_done 在主线程执行
        self._signals.done.connect(self._on_done)

    @Slot(str, QImage)
    def _on_done(self, key: str, img: QImage) -> None:
        # 主线程创建 QPixmap；失败时发空 QPixmap，由调用方判断
        pm = QPixmap() if img.isNull() else QPixmap.fromImage(img)
        self.ready.emit(key, pm)

    def request(self, key: str, path: Path, kind: str) -> None:
        task = _ThumbTask(key, path, kind, self._max_px, self._signals)
        self._pool.start(task)

    def shutdown(self) -> None:
        """退出时丢弃排队任务并等待在跑的任务结束，避免信号发往已销毁对象。"""
        self._pool.clear()
        self._pool.waitForDone(3000)
