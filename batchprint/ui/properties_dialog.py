"""单文档属性对话框：分页预览 + 单独打印设置（份数 / 页面范围 / 单独打印机设置）。

预览按需渲染当前页并缓存——一次性渲染整本 PDF 会重演"批量缩略图把主线程卡死"的问题。
"""
from __future__ import annotations

import logging

import fitz  # PyMuPDF
from PySide6.QtCore import Qt
from PySide6.QtGui import QImage, QImageReader, QPixmap
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from ..core import printers
from ..core.engine import parse_page_range
from ..core.settings import DocSettings, GlobalSettings, PrintJob, type_label

log = logging.getLogger(__name__)

PREVIEW_PX = 420


class PropertiesDialog(QDialog):
    """编辑单个任务的 DocSettings；左侧可逐页预览。确认后写回 job.doc。"""

    def __init__(self, job: PrintJob, global_settings: GlobalSettings, parent=None):
        super().__init__(parent)
        self.job = job
        self.global_settings = global_settings
        # 在副本上编辑，取消则不影响原对象
        self._devmode = job.doc.devmode

        self._page_cache: dict[int, QPixmap] = {}
        self._page_index = 0
        self._page_count = self._detect_page_count()

        self.setWindowTitle(f"属性 - {job.path.name}")
        self.setMinimumWidth(780)
        self._build_ui()
        self._load_from_job()
        self._show_page(0)

    # ---------------- 页数 / 渲染 ----------------
    def _detect_page_count(self) -> int:
        if self.job.kind != "pdf":
            return 1
        if self.job.pages:
            return self.job.pages
        try:
            with fitz.open(self.job.path) as doc:
                return doc.page_count
        except Exception as exc:  # noqa: BLE001
            log.warning("读取页数失败 %s: %s", self.job.path, exc)
            return 0

    def _render_page(self, idx: int) -> QPixmap | None:
        """按需渲染第 idx 页（0 基）并缓存。"""
        if idx in self._page_cache:
            return self._page_cache[idx]
        try:
            if self.job.kind == "pdf":
                with fitz.open(self.job.path) as doc:
                    if not (0 <= idx < doc.page_count):
                        return None
                    page = doc.load_page(idx)
                    rect = page.rect
                    scale = PREVIEW_PX / max(rect.width, rect.height, 1)
                    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                    img = QImage(pix.samples, pix.width, pix.height,
                                 pix.stride, QImage.Format_RGB888).copy()
                    pm = QPixmap.fromImage(img)
            else:
                reader = QImageReader(str(self.job.path))
                size = reader.size()
                if size.isValid() and (size.width() > PREVIEW_PX
                                       or size.height() > PREVIEW_PX):
                    reader.setScaledSize(
                        size.scaled(PREVIEW_PX, PREVIEW_PX, Qt.KeepAspectRatio)
                    )
                img = reader.read()
                if img.isNull():
                    return None
                pm = QPixmap.fromImage(img)
        except Exception as exc:  # noqa: BLE001
            log.warning("预览渲染失败 %s 第%d页: %s", self.job.path, idx + 1, exc)
            return None

        self._page_cache[idx] = pm
        return pm

    # ---------------- UI ----------------
    def _build_ui(self) -> None:
        root = QHBoxLayout(self)

        # ---- 左侧：分页预览 ----
        left = QVBoxLayout()
        self.preview_label = QLabel()
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setMinimumSize(PREVIEW_PX, PREVIEW_PX)
        self.preview_label.setStyleSheet(
            "border:1px solid #C6C6C6; background:#FAFAFA;"
        )
        left.addWidget(self.preview_label)

        # 翻页导航（图片只有一页时隐藏）
        self.nav = QWidget()
        nav_row = QHBoxLayout(self.nav)
        nav_row.setContentsMargins(0, 0, 0, 0)
        self.btn_prev = QPushButton("◀ 上一页")
        self.btn_next = QPushButton("下一页 ▶")
        self.spin_page = QSpinBox()
        self.spin_page.setRange(1, max(1, self._page_count))
        self.spin_page.setAlignment(Qt.AlignCenter)
        self.lbl_total = QLabel(f"/ {self._page_count} 页")
        self.btn_prev.clicked.connect(lambda: self._step_page(-1))
        self.btn_next.clicked.connect(lambda: self._step_page(1))
        self.spin_page.valueChanged.connect(lambda v: self._show_page(v - 1))
        nav_row.addWidget(self.btn_prev)
        nav_row.addStretch(1)
        nav_row.addWidget(QLabel("第"))
        nav_row.addWidget(self.spin_page)
        nav_row.addWidget(self.lbl_total)
        nav_row.addStretch(1)
        nav_row.addWidget(self.btn_next)
        left.addWidget(self.nav)
        self.nav.setVisible(self._page_count > 1)

        # 该页是否会被打印（受"页面范围"影响）
        self.lbl_in_range = QLabel("")
        self.lbl_in_range.setAlignment(Qt.AlignCenter)
        left.addWidget(self.lbl_in_range)

        info = QLabel(self._info_text())
        info.setWordWrap(True)
        info.setStyleSheet("color:#555;")
        left.addWidget(info)
        left.addStretch(1)
        root.addLayout(left)

        # ---- 右侧：设置 ----
        right = QVBoxLayout()
        form = QFormLayout()

        row_copies = QWidget()
        hc = QHBoxLayout(row_copies)
        hc.setContentsMargins(0, 0, 0, 0)
        self.chk_copies = QCheckBox("自定义份数")
        self.spin_copies = QSpinBox()
        self.spin_copies.setRange(1, 999)
        self.spin_copies.setValue(self.global_settings.copies)
        hc.addWidget(self.chk_copies)
        hc.addWidget(self.spin_copies)
        hc.addStretch(1)
        form.addRow("份数", row_copies)
        self.chk_copies.toggled.connect(self.spin_copies.setEnabled)

        self.edit_range = QLineEdit()
        self.edit_range.setPlaceholderText("留空=全部，例如 1-3,5")
        if self.job.kind != "pdf":
            self.edit_range.setEnabled(False)
            self.edit_range.setPlaceholderText("（仅 PDF 可用）")
        self.edit_range.textChanged.connect(self._refresh_range_hint)
        form.addRow("页面范围", self.edit_range)

        right.addLayout(form)

        self.lbl_devmode = QLabel()
        self.lbl_devmode.setStyleSheet("color:#333;")
        right.addWidget(self.lbl_devmode)

        btn_row = QHBoxLayout()
        self.btn_dev = QPushButton("调用打印机设置…")
        self.btn_dev.clicked.connect(self._on_open_devmode)
        self.btn_clear_dev = QPushButton("恢复继承全局")
        self.btn_clear_dev.clicked.connect(self._on_clear_devmode)
        btn_row.addWidget(self.btn_dev)
        btn_row.addWidget(self.btn_clear_dev)
        btn_row.addStretch(1)
        right.addLayout(btn_row)
        right.addStretch(1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        right.addWidget(buttons)

        root.addLayout(right, 1)

    def _info_text(self) -> str:
        parts = [f"文件：{self.job.path.name}", f"类型：{type_label(self.job)}"]
        if self.job.kind == "pdf":
            parts.append(f"页数：{self._page_count}")
        return "\n".join(parts)

    # ---------------- 翻页 ----------------
    def _step_page(self, delta: int) -> None:
        self._show_page(self._page_index + delta)

    def _show_page(self, idx: int) -> None:
        if self._page_count <= 0:
            self.preview_label.setText("（无法预览）")
            return
        idx = max(0, min(idx, self._page_count - 1))
        self._page_index = idx

        pm = self._render_page(idx)
        if pm is None:
            self.preview_label.setText("（无法预览此页）")
        else:
            self.preview_label.setPixmap(pm)

        self.spin_page.blockSignals(True)
        self.spin_page.setValue(idx + 1)
        self.spin_page.blockSignals(False)
        self.btn_prev.setEnabled(idx > 0)
        self.btn_next.setEnabled(idx < self._page_count - 1)
        self._refresh_range_hint()

    def _refresh_range_hint(self) -> None:
        """提示当前这页在不在「页面范围」里，即会不会被打印。"""
        if self.job.kind != "pdf" or self._page_count <= 0:
            self.lbl_in_range.setText("")
            return
        spec = self.edit_range.text().strip()
        if not spec:
            self.lbl_in_range.setText("")
            return
        pages = parse_page_range(spec, self._page_count)
        if self._page_index in pages:
            self.lbl_in_range.setText("本页在打印范围内")
            self.lbl_in_range.setStyleSheet("color:#16A34A;")
        else:
            self.lbl_in_range.setText("本页不在打印范围内，不会打印")
            self.lbl_in_range.setStyleSheet("color:#C2410C;")

    # ---------------- 数据装载/回写 ----------------
    def _load_from_job(self) -> None:
        doc = self.job.doc
        self.chk_copies.setChecked(doc.copies is not None)
        self.spin_copies.setEnabled(doc.copies is not None)
        if doc.copies is not None:
            self.spin_copies.setValue(doc.copies)
        if doc.page_range:
            self.edit_range.setText(doc.page_range)
        self._refresh_devmode_label()

    def _refresh_devmode_label(self) -> None:
        if self._devmode is not None:
            self.lbl_devmode.setText("打印机设置：已为本文档单独设置")
        else:
            self.lbl_devmode.setText("打印机设置：继承全局设置")

    # ---------------- 事件 ----------------
    def _on_open_devmode(self) -> None:
        name = self.global_settings.printer_name
        if not name:
            self.lbl_devmode.setText("打印机设置：请先在主界面选择打印机")
            return
        base = self._devmode if self._devmode is not None else self.global_settings.devmode
        dm = printers.open_printer_dialog(int(self.winId()), name, base)
        if dm is not None:
            self._devmode = dm
            self._refresh_devmode_label()

    def _on_clear_devmode(self) -> None:
        self._devmode = None
        self._refresh_devmode_label()

    def accept(self) -> None:
        self.job.doc = DocSettings(
            devmode=self._devmode,
            copies=self.spin_copies.value() if self.chk_copies.isChecked() else None,
            page_range=self.edit_range.text().strip() or None,
        )
        super().accept()
