"""主窗口：工具栏 + 打印机/全局设置行 + 文件列表表格 + 进度/控制。"""
from __future__ import annotations

import logging
from pathlib import Path

from PySide6.QtCore import QSize, Qt, QTimer
from PySide6.QtGui import QColor, QIcon, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

import os

from ..core import importer, printers
from ..core.converter import NO_CONVERTER_MSG, WordConverter
from ..core.runner import JobRunner
from ..core.scanner import ScanWorker
from ..core.settings import (
    FIT_CENTER,
    FIT_FILL,
    FIT_STRETCH,
    STATUS_CONVERT_FAILED,
    STATUS_CONVERTING,
    STATUS_FAILED,
    STATUS_PENDING,
    STATUS_PRINTING,
    GlobalSettings,
    PrintJob,
    effective_markup,
    type_label,
)
from ..core.thumbnailer import Thumbnailer
from ..resources import app_icon_path
from .properties_dialog import PropertiesDialog

log = logging.getLogger(__name__)

COL_CHECK = 0
COL_THUMB = 1
COL_NAME = 2
COL_TYPE = 3
COL_PAGES = 4
COL_COPIES = 5
COL_STATUS = 6
COL_PROPS = 7
COL_COUNT = 8

THUMB_PX = 64
ROW_HEIGHT = THUMB_PX + 8
INSERT_CHUNK = 20        # 每批插入的行数，批间让出事件循环保持界面响应
THUMB_PREFETCH = 5       # 可见范围上下额外预取的行数，滚动时更跟手

FAILED_STATUSES = (STATUS_FAILED, STATUS_CONVERT_FAILED)
FAILED_COLOR = QColor("#DC2626")    # 失败状态用红字

FIT_LABELS = [
    ("等比缩放居中", FIT_CENTER),
    ("等比铺满(裁切)", FIT_FILL),
    ("拉伸铺满", FIT_STRETCH),
]

# 主要操作按钮（添加文件）—— 蓝色，与程序图标同色系
PRIMARY_BTN_QSS = """
QPushButton {
    background-color: #2F6BFF;
    color: #FFFFFF;
    border: none;
    border-radius: 6px;
    padding: 7px 18px;
    font-weight: bold;
}
QPushButton:hover   { background-color: #1E56E0; }
QPushButton:pressed { background-color: #1846BC; }
QPushButton:disabled{ background-color: #BCC9E8; color: #EFEFEF; }
"""

# 执行按钮（开始打印）—— 绿色，表示"开始/执行"
ACTION_BTN_QSS = """
QPushButton {
    background-color: #16A34A;
    color: #FFFFFF;
    border: none;
    border-radius: 6px;
    padding: 7px 18px;
    font-weight: bold;
}
QPushButton:hover   { background-color: #12873D; }
QPushButton:pressed { background-color: #0E6E32; }
QPushButton:disabled{ background-color: #BEDCC8; color: #EFEFEF; }
"""

# 全局设置分组框：灰色细边框 + 灰色标题
GROUP_BOX_QSS = """
QGroupBox {
    border: 1px solid #C6C6C6;
    border-radius: 6px;
    margin-top: 10px;
    padding: 10px 8px 8px 8px;
    font-weight: bold;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 10px;
    padding: 0 5px;
    color: #666666;
}
"""


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("批量打印")
        self.resize(920, 620)
        _icon = app_icon_path()
        if os.path.exists(_icon):
            self.setWindowIcon(QIcon(_icon))

        self.jobs: list[PrintJob] = []
        self.global_settings = GlobalSettings()
        self.runner: JobRunner | None = None

        # 导入相关状态（后台扫描 + 分批插入 + 缩略图进度）
        self.scanner: ScanWorker | None = None
        self._pending_jobs: list[PrintJob] = []
        self._total_to_add = 0
        self._added_count = 0
        self._thumb_total = 0
        self._thumb_pending = 0
        self._thumb_requested: set[str] = set()

        self.thumbnailer = Thumbnailer(max_px=THUMB_PX)
        self.thumbnailer.ready.connect(self._on_thumb_ready)

        # Word 文档转 PDF：唯一的后台转换线程，不会出现 Word 窗口
        self.converter = WordConverter(self)
        self.converter.converted.connect(self._on_word_converted)
        self.converter.failed.connect(self._on_word_failed)
        self.converter.start()
        self._conv_waiting: set[str] = set()   # 已排队、还没出结果的后台转换
        self._conv_total = 0

        self._build_ui()
        self._load_printers()
        self._update_count_label()
        self.setAcceptDrops(True)

    # ---------------- UI 构建 ----------------
    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)

        # 工具栏行
        tool_row = QHBoxLayout()
        self.btn_add_files = QPushButton("添加文件")
        self.btn_add_folder = QPushButton("添加文件夹")
        self.btn_remove = QPushButton("移除选中")
        self.btn_clear = QPushButton("清空")
        self.btn_check_all = QPushButton("全选")
        self.btn_uncheck_all = QPushButton("全不选")
        for b in (self.btn_add_files, self.btn_add_folder, self.btn_remove,
                  self.btn_clear, self.btn_check_all, self.btn_uncheck_all):
            tool_row.addWidget(b)
        tool_row.addStretch(1)
        # 常驻的勾选统计（放工具栏，避免被状态栏的临时消息覆盖）
        self.lbl_count = QLabel()
        self.lbl_count.setStyleSheet("color:#333;")
        tool_row.addWidget(self.lbl_count)
        root.addLayout(tool_row)

        # 「添加文件」是入口主操作，突出显示
        self.btn_add_files.setStyleSheet(PRIMARY_BTN_QSS)
        self.btn_add_files.setCursor(Qt.PointingHandCursor)

        self.btn_add_files.clicked.connect(self._on_add_files)
        self.btn_add_folder.clicked.connect(self._on_add_folder)
        self.btn_remove.clicked.connect(self._on_remove_selected)
        self.btn_clear.clicked.connect(self._on_clear)
        self.btn_check_all.clicked.connect(lambda: self._set_all_checked(True))
        self.btn_uncheck_all.clicked.connect(lambda: self._set_all_checked(False))

        # 全局设置（加框 + 说明，明确区分「全局」与「单文档」）
        global_box = QGroupBox("全局设置")
        global_box.setStyleSheet(GROUP_BOX_QSS)
        gb_layout = QVBoxLayout(global_box)

        set_row = QHBoxLayout()
        set_row.addWidget(QLabel("打印机:"))
        self.cmb_printer = QComboBox()
        self.cmb_printer.setMinimumWidth(240)
        self.cmb_printer.currentTextChanged.connect(self._on_printer_changed)
        set_row.addWidget(self.cmb_printer)

        self.btn_printer_setup = QPushButton("打印机设置…")
        self.btn_printer_setup.clicked.connect(self._on_global_printer_setup)
        set_row.addWidget(self.btn_printer_setup)

        set_row.addSpacing(16)
        set_row.addWidget(QLabel("份数:"))
        self.spin_copies = QSpinBox()
        self.spin_copies.setRange(1, 999)
        self.spin_copies.setValue(1)
        self.spin_copies.valueChanged.connect(self._on_copies_changed)
        set_row.addWidget(self.spin_copies)

        set_row.addSpacing(16)
        set_row.addWidget(QLabel("图片排版:"))
        self.cmb_fit = QComboBox()
        for label, _ in FIT_LABELS:
            self.cmb_fit.addItem(label)
        self.cmb_fit.currentIndexChanged.connect(self._on_fit_changed)
        set_row.addWidget(self.cmb_fit)

        set_row.addSpacing(16)
        self.chk_auto_orient = QCheckBox("自动匹配文档方向")
        self.chk_auto_orient.setChecked(self.global_settings.auto_orient)
        self.chk_auto_orient.setToolTip(
            "勾选后：横向的页面/图片自动用横向打印，纵向的用纵向（PDF、Word 逐页判断）。"
        )
        self.chk_auto_orient.toggled.connect(self._on_auto_orient_changed)
        set_row.addWidget(self.chk_auto_orient)

        set_row.addStretch(1)
        gb_layout.addLayout(set_row)

        hint = QLabel(
            "以上设置应用到全部文档。如需单独设置某个文档：可在下方列表的「份数」列"
            "行内直接修改，或点击该行右侧的「属性」按钮（可设份数、页面范围、"
            "以及该文档单独的打印机设置）。"
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color:#666;")
        gb_layout.addWidget(hint)

        root.addWidget(global_box)

        # 表格
        self.table = QTableWidget(0, COL_COUNT)
        self.table.setHorizontalHeaderLabels(
            ["", "预览", "文件名", "类型", "页数", "份数", "状态", "属性"]
        )
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setIconSize(QSize(THUMB_PX, THUMB_PX))
        self.table.itemChanged.connect(self._on_item_changed)
        # 滚动到哪，就给哪些行生成缩略图
        self.table.verticalScrollBar().valueChanged.connect(
            self._request_visible_thumbs
        )

        # 注意：除文件名外一律用固定宽度，不要用 ResizeToContents——
        # 后者每次缩略图到位都会重新测量所有行，行一多就是 O(行数²) 的卡顿。
        header = self.table.horizontalHeader()
        for col, width in (
            (COL_CHECK, 34), (COL_THUMB, THUMB_PX + 12), (COL_TYPE, 92),
            (COL_PAGES, 60), (COL_COPIES, 76), (COL_STATUS, 72), (COL_PROPS, 62),
        ):
            header.setSectionResizeMode(col, QHeaderView.Interactive)
            self.table.setColumnWidth(col, width)
        header.setSectionResizeMode(COL_NAME, QHeaderView.Stretch)
        root.addWidget(self.table, 1)

        # 进度/控制行
        ctrl_row = QHBoxLayout()
        self.progress = QProgressBar()
        self.progress.setTextVisible(True)
        self.lbl_status = QLabel("就绪")
        self.btn_start = QPushButton("开始打印")
        self.btn_stop = QPushButton("停止")
        self.btn_stop.setEnabled(False)
        self.btn_start.clicked.connect(self._on_start)
        self.btn_stop.clicked.connect(self._on_stop)
        # 「开始打印」是执行主操作，突出显示
        self.btn_start.setStyleSheet(ACTION_BTN_QSS)
        self.btn_start.setCursor(Qt.PointingHandCursor)
        ctrl_row.addWidget(self.progress, 1)
        ctrl_row.addWidget(self.lbl_status)
        ctrl_row.addWidget(self.btn_start)
        ctrl_row.addWidget(self.btn_stop)
        root.addLayout(ctrl_row)

    # ---------------- 打印机 ----------------
    def _load_printers(self) -> None:
        names = printers.list_printers()
        self.cmb_printer.blockSignals(True)
        self.cmb_printer.clear()
        self.cmb_printer.addItems(names)
        default = printers.default_printer()
        if default and default in names:
            self.cmb_printer.setCurrentText(default)
        self.cmb_printer.blockSignals(False)
        self.global_settings.printer_name = self.cmb_printer.currentText()
        if not names:
            self.lbl_status.setText("未检测到打印机")

    def _on_printer_changed(self, name: str) -> None:
        self.global_settings.printer_name = name
        # 换打印机后，旧的 DEVMODE 不再适用：清空全局与各文档的单独设置
        self.global_settings.devmode = None
        for job in self.jobs:
            job.doc.devmode = None
        self.lbl_status.setText(f"已选择打印机：{name}（打印机设置已重置）")

    def _on_global_printer_setup(self) -> None:
        name = self.global_settings.printer_name
        if not name:
            QMessageBox.information(self, "提示", "请先选择一台打印机。")
            return
        dm = printers.open_printer_dialog(
            int(self.winId()), name, self.global_settings.devmode
        )
        if dm is not None:
            self.global_settings.devmode = dm
            self.lbl_status.setText("全局打印机设置已更新，将应用到所有文档。")

    def _on_copies_changed(self, value: int) -> None:
        self.global_settings.copies = value
        # 只有"继承全局"的行跟着变；屏蔽信号，避免把它们误标记成单独设置
        for row in range(self.table.rowCount()):
            if self.jobs[row].doc.copies is None:
                spin = self.table.cellWidget(row, COL_COPIES)
                if isinstance(spin, QSpinBox):
                    spin.blockSignals(True)
                    spin.setValue(value)
                    spin.blockSignals(False)
        self._update_count_label()

    def _on_fit_changed(self, index: int) -> None:
        self.global_settings.image_fit = FIT_LABELS[index][1]

    def _on_auto_orient_changed(self, checked: bool) -> None:
        self.global_settings.auto_orient = checked

    # ---------------- 文件导入 ----------------
    def _on_add_files(self) -> None:
        exts = " ".join(f"*{e}" for e in sorted(importer.SUPPORTED_EXTS))
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择文件", "", f"支持的文件 ({exts})"
        )
        if files:
            self._add_paths(files)

    def _on_add_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            self._add_paths([folder])

    def _add_paths(self, paths: list[str]) -> None:
        """异步导入：后台扫描 -> 分批插入行，全程显示「正在加载 X/Y」。"""
        if self.scanner and self.scanner.isRunning():
            self.lbl_status.setText("正在加载中，请稍候…")
            return
        existing = {str(j.path.resolve()).lower() for j in self.jobs}
        self._set_import_busy(True)
        self.lbl_status.setText("正在扫描文件…")
        self.progress.setRange(0, 0)          # 未知总数时显示忙碌态
        self.scanner = ScanWorker(paths, existing, self)
        self.scanner.progress.connect(self._on_scan_progress)
        self.scanner.done.connect(self._on_scan_done)
        self.scanner.start()

    def _on_scan_progress(self, done: int, total: int) -> None:
        self.progress.setRange(0, max(1, total))
        self.progress.setValue(done)
        self.lbl_status.setText(f"正在扫描… {done}/{total}")

    def _on_scan_done(self, jobs: list) -> None:
        if not jobs:
            self._set_import_busy(False)
            self.progress.setRange(0, 1)
            self.progress.setValue(0)
            self.lbl_status.setText("没有新的可打印文件被添加。")
            return
        self._pending_jobs = list(jobs)
        self._total_to_add = len(jobs)
        self._added_count = 0
        self.progress.setRange(0, self._total_to_add)
        self._insert_next_chunk()

    def _insert_next_chunk(self) -> None:
        """分批插入行，每批之间让出事件循环，界面始终可响应。"""
        chunk = self._pending_jobs[:INSERT_CHUNK]
        del self._pending_jobs[:INSERT_CHUNK]

        self.table.setUpdatesEnabled(False)
        try:
            for job in chunk:
                self._append_job_row(job)
        finally:
            self.table.setUpdatesEnabled(True)

        self._added_count += len(chunk)
        self.progress.setValue(self._added_count)
        self.lbl_status.setText(f"正在加载… {self._added_count}/{self._total_to_add}")
        self._request_visible_thumbs()
        self._update_count_label()

        if self._pending_jobs:
            QTimer.singleShot(0, self._insert_next_chunk)   # 让出事件循环
        else:
            self._set_import_busy(False)
            # 进度条归位，避免导入完还停在满格、看着像有任务在跑
            self.progress.setRange(0, 1)
            self.progress.setValue(0)
            self.lbl_status.setText(
                f"已添加 {self._total_to_add} 个文件，共 {len(self.jobs)} 个。"
            )
            self._update_thumb_status()

    def _set_import_busy(self, busy: bool) -> None:
        for w in (self.btn_add_files, self.btn_add_folder,
                  self.btn_remove, self.btn_clear, self.btn_start):
            w.setEnabled(not busy)

    def _append_job_row(self, job: PrintJob) -> None:
        if job.kind == "word":
            self._queue_conversion(job)
        self.jobs.append(job)
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setRowHeight(row, ROW_HEIGHT)

        self.table.blockSignals(True)

        chk = QTableWidgetItem()
        chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        chk.setCheckState(Qt.Checked if job.checked else Qt.Unchecked)
        self.table.setItem(row, COL_CHECK, chk)

        thumb = QTableWidgetItem()
        thumb.setTextAlignment(Qt.AlignCenter)
        self.table.setItem(row, COL_THUMB, thumb)

        name = QTableWidgetItem(job.path.name)
        name.setToolTip(str(job.path))
        self.table.setItem(row, COL_NAME, name)

        self.table.setItem(row, COL_TYPE, QTableWidgetItem(type_label(job)))
        if job.pages is not None:
            pages_text = str(job.pages)
        else:   # Word 文档转换完成前还不知道页数
            pages_text = "…" if job.status == STATUS_CONVERTING else "-"
        self.table.setItem(row, COL_PAGES, QTableWidgetItem(pages_text))
        spin = QSpinBox()
        spin.setRange(1, 999)
        spin.setValue(self._effective_copies(job))
        spin.setAlignment(Qt.AlignCenter)
        spin.valueChanged.connect(
            lambda v, j=job, s=spin: self._on_row_copies_changed(j, s, v)
        )
        self._style_copies_spin(spin, job)
        self.table.setCellWidget(row, COL_COPIES, spin)

        self.table.setItem(row, COL_STATUS, QTableWidgetItem())
        self._set_status(row, job.status, job.convert_error or "")

        btn = QPushButton("属性")
        btn.clicked.connect(lambda _=False, j=job: self._open_properties(j))
        self.table.setCellWidget(row, COL_PROPS, btn)

        self.table.blockSignals(False)

    def _set_status(self, row: int, text: str, tooltip: str | None = None) -> None:
        """更新状态列。失败类状态用红字；tooltip 为 None 表示保留原提示。"""
        item = self.table.item(row, COL_STATUS)
        if item is None:
            return
        item.setText(text)
        if tooltip is not None:
            item.setToolTip(tooltip)
        if text in FAILED_STATUSES:
            item.setForeground(FAILED_COLOR)
        else:
            item.setData(Qt.ForegroundRole, None)      # 恢复默认颜色

    def _set_row_checked(self, row: int, checked: bool) -> None:
        self.jobs[row].checked = checked
        self.table.blockSignals(True)
        self.table.item(row, COL_CHECK).setCheckState(Qt.Checked if checked else Qt.Unchecked)
        self.table.blockSignals(False)

    # ---------------- Word 文档转换 ----------------
    def _queue_conversion(self, job: PrintJob) -> None:
        """导入后立即在后台排队转成 PDF；转完补上页数和缩略图。"""
        if not self.converter.available:
            job.status = STATUS_CONVERT_FAILED
            job.convert_error = NO_CONVERTER_MSG
            job.checked = False              # 打不出来，默认不勾选
            return
        job.status = STATUS_CONVERTING
        key = str(job.path)
        self._conv_waiting.add(key)
        self._conv_total += 1
        self.converter.request(key, job.path, effective_markup(self.global_settings, job))

    def _on_word_converted(self, key: str, pdf: str, pages: int, markup: bool) -> None:
        self._conv_waiting.discard(key)
        row = self._row_of_key(key)
        if row >= 0:
            job = self.jobs[row]
            # 带修订 / 不带修订是两个不同的 PDF，只采用和当前设置一致的那个
            if markup == effective_markup(self.global_settings, job):
                job.pdf_path = Path(pdf)
                job.pages = pages
                job.convert_error = None
                self.table.item(row, COL_PAGES).setText(str(pages))
                if job.status in (STATUS_CONVERTING, STATUS_CONVERT_FAILED):
                    job.status = STATUS_PENDING
                    self._set_status(row, STATUS_PENDING, "")
                self._update_count_label()
                self._request_visible_thumbs()
        self._update_thumb_status()

    def _on_word_failed(self, key: str, msg: str) -> None:
        was_waiting = key in self._conv_waiting
        self._conv_waiting.discard(key)
        row = self._row_of_key(key)
        # 打印时的转换失败由打印线程自己标记，这里只处理导入后的后台转换
        if was_waiting and row >= 0:
            job = self.jobs[row]
            job.convert_error = msg
            if job.status == STATUS_CONVERTING:
                job.status = STATUS_CONVERT_FAILED
                self._set_row_checked(row, False)   # 打不出来，自动取消勾选
                self._update_count_label()
            self.table.item(row, COL_PAGES).setText("-")
            self._set_status(row, job.status, msg)
        self._update_thumb_status()

    # ---------------- 缩略图：只做屏幕上看得见的那些行 ----------------
    def _request_visible_thumbs(self) -> None:
        """只为可见（及上下预取几行）的行生成缩略图。

        一次性给上百个文件渲染缩略图会让 Python 工作线程长时间持有 GIL，
        把主线程饿死；而用户一次也只能看到十来行。
        """
        rows = self.table.rowCount()
        if rows == 0:
            return
        first = self.table.rowAt(0)
        last = self.table.rowAt(self.table.viewport().height() - 1)
        first = 0 if first < 0 else first
        last = rows - 1 if last < 0 else last
        first = max(0, first - THUMB_PREFETCH)
        last = min(rows - 1, last + THUMB_PREFETCH)

        for row in range(first, last + 1):
            job = self.jobs[row]
            key = str(job.path)
            if key in self._thumb_requested:
                continue
            if job.kind == "word":
                if job.pdf_path is None:
                    continue        # 还没转好；转换完成后会再次触发这里
                src, kind = job.pdf_path, "pdf"
            else:
                src, kind = job.path, job.kind
            self._thumb_requested.add(key)
            self._thumb_total += 1
            self._thumb_pending += 1
            self.thumbnailer.request(key, src, kind)

    def _row_of_key(self, key: str) -> int:
        for i, job in enumerate(self.jobs):
            if str(job.path) == key:
                return i
        return -1

    def _on_thumb_ready(self, key: str, pixmap: QPixmap) -> None:
        self._thumb_pending = max(0, self._thumb_pending - 1)
        if not pixmap.isNull():
            row = self._row_of_key(key)
            if row >= 0:
                item = self.table.item(row, COL_THUMB)
                if item is not None:
                    item.setIcon(QIcon(pixmap))
        self._update_thumb_status()

    def _update_thumb_status(self) -> None:
        """Word 转换、缩略图都是异步陆续完成的，这里给个进度，别让用户以为卡住了。"""
        if self.scanner and self.scanner.isRunning():
            return                      # 扫描阶段的提示优先
        if self._pending_jobs:
            return                      # 还在插入行
        if self.runner and self.runner.isRunning():
            return                      # 打印中，状态栏显示打印进度
        if self._conv_waiting:
            done = self._conv_total - len(self._conv_waiting)
            self.lbl_status.setText(f"正在转换 Word 文档… {done}/{self._conv_total}")
        elif self._thumb_pending > 0:
            done = self._thumb_total - self._thumb_pending
            self.lbl_status.setText(f"正在生成预览… {done}/{self._thumb_total}")
        elif self._thumb_total or self._conv_total:
            self.lbl_status.setText(f"就绪，共 {len(self.jobs)} 个文件。")
            self._thumb_total = 0
            self._conv_total = 0

    def _open_properties(self, job: PrintJob) -> None:
        dlg = PropertiesDialog(job, self.global_settings, self)
        if dlg.exec():
            self._refresh_copies_cell(job)

    # ---------------- 份数（全局 / 单文档） ----------------
    def _effective_copies(self, job: PrintJob) -> int:
        """该文档实际会打的份数：有单独设置用单独的，否则跟随全局。"""
        if job.doc.copies is not None:
            return job.doc.copies
        return self.global_settings.copies

    def _style_copies_spin(self, spin: QSpinBox, job: PrintJob) -> None:
        if job.doc.copies is not None:
            spin.setStyleSheet("font-weight:bold;")
            spin.setToolTip("已为本文档单独设置份数（不跟随全局）。\n可在「属性」里取消勾选恢复继承。")
        else:
            spin.setStyleSheet("")
            spin.setToolTip("正跟随顶部的全局份数。\n直接在这里修改即可为本文档单独设置。")

    def _on_row_copies_changed(self, job: PrintJob, spin: QSpinBox, value: int) -> None:
        """行内改份数 = 为该文档设定单独份数。"""
        job.doc.copies = value
        self._style_copies_spin(spin, job)
        self._update_count_label()

    def _refresh_copies_cell(self, job: PrintJob) -> None:
        """属性窗口改过之后，把行内份数同步回来。"""
        try:
            row = self.jobs.index(job)
        except ValueError:
            return
        spin = self.table.cellWidget(row, COL_COPIES)
        if isinstance(spin, QSpinBox):
            spin.blockSignals(True)
            spin.setValue(self._effective_copies(job))
            spin.blockSignals(False)
            self._style_copies_spin(spin, job)
        self._update_count_label()

    # ---------------- 列表操作 ----------------
    def _on_remove_selected(self) -> None:
        rows = sorted({i.row() for i in self.table.selectedIndexes()}, reverse=True)
        if not rows:
            self.lbl_status.setText("请先选中要移除的行。")
            return
        for row in rows:
            self._conv_waiting.discard(str(self.jobs[row].path))
            self.table.removeRow(row)
            del self.jobs[row]
        self._rebind_property_buttons()
        self._update_count_label()
        self.lbl_status.setText(f"已移除 {len(rows)} 项，共 {len(self.jobs)} 个。")

    def _on_clear(self) -> None:
        self.table.setRowCount(0)
        self.jobs.clear()
        self._pending_jobs.clear()
        self._thumb_total = 0
        self._thumb_pending = 0
        self._thumb_requested.clear()
        self.converter.cancel_pending()
        self._conv_waiting.clear()
        self._conv_total = 0
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        self._update_count_label()
        self.lbl_status.setText("已清空。")

    def _rebind_property_buttons(self) -> None:
        """删除行后行号变化，重新绑定属性按钮与缩略图行号。"""
        for row in range(self.table.rowCount()):
            job = self.jobs[row]
            btn = self.table.cellWidget(row, COL_PROPS)
            if isinstance(btn, QPushButton):
                try:
                    btn.clicked.disconnect()
                except (RuntimeError, TypeError):
                    pass
                btn.clicked.connect(lambda _=False, j=job: self._open_properties(j))

    def _set_all_checked(self, checked: bool) -> None:
        self.table.blockSignals(True)
        state = Qt.Checked if checked else Qt.Unchecked
        for row in range(self.table.rowCount()):
            self.table.item(row, COL_CHECK).setCheckState(state)
            self.jobs[row].checked = checked
        self.table.blockSignals(False)
        self._update_count_label()

    def _on_item_changed(self, item: QTableWidgetItem) -> None:
        if item.column() == COL_CHECK:
            row = item.row()
            if 0 <= row < len(self.jobs):
                self.jobs[row].checked = item.checkState() == Qt.Checked
                self._update_count_label()

    def _update_count_label(self) -> None:
        """常驻显示：勾选了几个 / 共几个（勾选的才会打印）。"""
        total = len(self.jobs)
        if total == 0:
            self.lbl_count.setText("尚未添加文件")
            return
        checked = sum(1 for j in self.jobs if j.checked)
        # 算上份数才是真实出纸量（图片按 1 页算）
        pages = sum((j.pages or 1) * self._effective_copies(j)
                    for j in self.jobs if j.checked)
        self.lbl_count.setText(
            f"已勾选 {checked} / {total} 个文件（约 {pages} 页）"
        )

    # ---------------- 拖拽 ----------------
    def dragEnterEvent(self, event) -> None:  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event) -> None:  # noqa: N802
        paths = [u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()]
        if paths:
            self._add_paths(paths)

    # ---------------- 打印控制 ----------------
    def _on_start(self) -> None:
        if self.runner and self.runner.isRunning():
            return
        if not self.global_settings.printer_name:
            QMessageBox.information(self, "提示", "请先选择一台打印机。")
            return
        selected = [(row, job) for row, job in enumerate(self.jobs) if job.checked]
        if not selected:
            QMessageBox.information(self, "提示", "请至少勾选一个要打印的文件。")
            return

        # 重置选中项状态
        self.table.blockSignals(True)
        for row, job in selected:
            job.status = STATUS_PENDING
            job.error = None
            self._set_status(row, STATUS_PENDING, "")
        self.table.blockSignals(False)

        self.progress.setRange(0, len(selected))
        self.progress.setValue(0)
        self._set_controls_enabled(False)

        self.runner = JobRunner(self.global_settings, selected, self.converter, self)
        self.runner.job_started.connect(self._on_job_started)
        self.runner.job_finished.connect(self._on_job_finished)
        self.runner.progress.connect(self._on_progress)
        self.runner.finished_all.connect(self._on_finished_all)
        self.runner.start()
        self.lbl_status.setText(f"开始打印，共 {len(selected)} 个文件…")

    def _on_stop(self) -> None:
        if self.runner and self.runner.isRunning():
            self.runner.stop()
            self.lbl_status.setText("正在停止…（完成当前文件后停止）")
            self.btn_stop.setEnabled(False)

    def _on_job_started(self, row: int) -> None:
        if 0 <= row < self.table.rowCount():
            self.jobs[row].status = STATUS_PRINTING
            self._set_status(row, STATUS_PRINTING)
            self.lbl_status.setText(f"正在打印：{self.jobs[row].path.name}")

    def _on_job_finished(self, row: int, ok: bool, msg: str) -> None:
        if 0 <= row < self.table.rowCount():
            self._set_status(row, self.jobs[row].status, "" if ok else msg)

    def _on_progress(self, done: int, total: int) -> None:
        self.progress.setValue(done)

    def _on_finished_all(self, done: int, failed: int) -> None:
        self._set_controls_enabled(True)
        self.btn_stop.setEnabled(False)
        if failed:
            self.lbl_status.setText(f"打印结束：完成 {done - failed}，失败 {failed}。")
        else:
            self.lbl_status.setText(f"打印结束：共完成 {done} 个文件。")

    def _set_controls_enabled(self, enabled: bool) -> None:
        for w in (self.btn_add_files, self.btn_add_folder, self.btn_remove,
                  self.btn_clear, self.btn_check_all, self.btn_uncheck_all,
                  self.cmb_printer, self.btn_printer_setup, self.spin_copies,
                  self.cmb_fit, self.chk_auto_orient, self.btn_start, self.table):
            w.setEnabled(enabled)
        self.btn_stop.setEnabled(not enabled)

    def closeEvent(self, event) -> None:  # noqa: N802
        if self.runner and self.runner.isRunning():
            self.runner.stop()
            self.runner.wait(3000)
        if self.scanner and self.scanner.isRunning():
            self.scanner.stop()
            self.scanner.wait(3000)
        self.converter.shutdown()           # 关闭后台 Word
        self.thumbnailer.shutdown()
        super().closeEvent(event)


def run() -> None:
    import sys
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
