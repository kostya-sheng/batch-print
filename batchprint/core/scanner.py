"""后台扫描：展开文件夹、读取 PDF 页数，并上报进度。

这些都是 I/O，放在主线程会让界面"未响应"（文件夹大或在网络盘时尤其明显），
因此统一挪到后台线程，主线程只负责显示「正在加载 X/Y」。
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Iterable

from PySide6.QtCore import QThread, Signal

from . import importer
from .settings import PrintJob

log = logging.getLogger(__name__)


class ScanWorker(QThread):
    """扫描路径 -> 去重 -> 逐个构建 PrintJob。

    信号：
      progress(done, total)  已处理/总数
      done(list[PrintJob])   扫描结果（被中止时是已完成的部分）
    """

    progress = Signal(int, int)
    done = Signal(list)

    def __init__(self, paths: Iterable[str | Path],
                 existing_keys: Iterable[str], parent=None):
        super().__init__(parent)
        self._paths = list(paths)
        self._existing = set(existing_keys)
        self._stop = False

    def stop(self) -> None:
        self._stop = True

    def run(self) -> None:  # noqa: D401 - QThread 入口
        try:
            files = importer.expand_paths(self._paths)
        except Exception:  # noqa: BLE001
            log.exception("扫描路径失败: %s", self._paths)
            files = []

        # 排除已在列表中的文件
        files = [f for f in files
                 if str(f.resolve()).lower() not in self._existing]

        total = len(files)
        self.progress.emit(0, total)

        jobs: list[PrintJob] = []
        step = max(1, total // 50)   # 控制信号频率，避免刷屏
        for i, path in enumerate(files, 1):
            if self._stop:
                break
            try:
                jobs.append(importer.build_job(path))
            except Exception:  # noqa: BLE001
                log.exception("读取文件信息失败: %s", path)
            if i % step == 0 or i == total:
                self.progress.emit(i, total)

        self.done.emit(jobs)
