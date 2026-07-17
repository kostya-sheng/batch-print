"""任务调度：在后台线程里按队列顺序串行打印，上报进度、支持中止。"""
from __future__ import annotations

import logging

from PySide6.QtCore import QThread, Signal

from . import engine
from .settings import (
    STATUS_DONE,
    STATUS_FAILED,
    GlobalSettings,
    PrintJob,
    resolve,
)

log = logging.getLogger(__name__)


class JobRunner(QThread):
    """串行执行选中的打印任务。

    信号：
      job_started(index)          某任务开始
      job_finished(index, ok, msg)某任务结束（ok=是否成功，msg=错误信息）
      progress(done, total)       已处理/总数
      finished_all(done, failed)  全部结束
    """

    job_started = Signal(int)
    job_finished = Signal(int, bool, str)
    progress = Signal(int, int)
    finished_all = Signal(int, int)

    def __init__(self, global_settings: GlobalSettings,
                 jobs: list[tuple[int, PrintJob]], parent=None):
        super().__init__(parent)
        self._global = global_settings
        self._jobs = jobs               # [(原始行号, job), ...]
        self._stop = False

    def stop(self) -> None:
        self._stop = True

    def _cancelled(self) -> bool:
        return self._stop

    def run(self) -> None:  # noqa: D401 - QThread 入口
        total = len(self._jobs)
        done = 0
        failed = 0
        for index, job in self._jobs:
            if self._stop:
                break
            self.job_started.emit(index)
            try:
                eff = resolve(self._global, job)
                engine.print_job(eff, job.path, job.kind, cancel=self._cancelled)
                job.status = STATUS_DONE
                job.error = None
                self.job_finished.emit(index, True, "")
            except engine.PrintCancelled:
                log.info("打印被中止于: %s", job.path)
                break
            except Exception as exc:  # noqa: BLE001
                failed += 1
                job.status = STATUS_FAILED
                job.error = str(exc)
                log.exception("打印失败: %s", job.path)
                self.job_finished.emit(index, False, str(exc))
            finally:
                done += 1
                self.progress.emit(done, total)

        self.finished_all.emit(done, failed)
