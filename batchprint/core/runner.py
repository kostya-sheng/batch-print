"""任务调度：在后台线程里按队列顺序串行打印，上报进度、支持中止。"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

from PySide6.QtCore import QThread, Signal

from . import engine
from .converter import (
    NO_CONVERTER_MSG,
    ConversionCancelled,
    ConversionError,
    WordConverter,
)
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
                 jobs: list[tuple[int, PrintJob]],
                 converter: Optional[WordConverter] = None, parent=None):
        super().__init__(parent)
        self._global = global_settings
        self._jobs = jobs               # [(原始行号, job), ...]
        self._converter = converter     # Word 文档转 PDF 用
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
                if job.kind == "word":
                    self._print_word(job, eff)
                else:
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
                if isinstance(exc, ConversionError):   # 文档本身的问题，原因已说清，不需要堆栈
                    log.warning("打印失败: %s: %s", job.path, exc)
                else:
                    log.exception("打印失败: %s", job.path)
                self.job_finished.emit(index, False, str(exc))
            finally:
                done += 1
                self.progress.emit(done, total)

        self.finished_all.emit(done, failed)

    def _print_word(self, job: PrintJob, eff) -> None:
        """Word 文档：先拿到转换好的 PDF（没转完就插队转换并等待），再按 PDF 打印。

        之前转换失败的也会在这里重试一次——比如上次只是 Word 临时卡住。
        """
        if self._converter is None or not self._converter.available:
            raise RuntimeError(NO_CONVERTER_MSG)
        try:
            pdf, _pages = self._converter.convert_blocking(
                str(job.path), job.path, eff.print_markup, self._cancelled)
        except ConversionCancelled:
            raise engine.PrintCancelled() from None
        engine.print_job(eff, Path(pdf), "pdf", cancel=self._cancelled,
                         doc_name=job.path.name)
