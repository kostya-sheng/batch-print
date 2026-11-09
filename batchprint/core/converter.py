"""Word 文档（.docx / .doc）→ PDF 转换：后台串行、隐藏 Word，结果缓存。

方案见开发蓝图第 9 节。要点：
- Word 的 COM 接口是单线程的：整个程序只有一个转换线程，逐个转换，复用同一个 Word 实例。
- 绝不让 Word 弹出窗口：隐藏、关闭提示、禁用宏、只读打开；传假密码，有密码的文档直接报错返回。
- 卡死保护：单个文件超时，就结束本程序启动的那个 Word 进程。只认启动时记下的 PID，
  不碰用户自己打开的 Word / WPS。
- 网络盘上的文件先复制到本地临时目录再转换。
- 转换结果按「路径 + 修改时间 + 大小 + 是否带修订」缓存，原文件一改就自动重新转换。
"""
from __future__ import annotations

import hashlib
import itertools
import logging
import os
import queue
import shutil
import tempfile
import threading
import time
import uuid
import winreg
import zipfile
from pathlib import Path
from typing import Callable, Optional

import fitz  # PyMuPDF
import pythoncom
import pywintypes
import win32api
import win32com.client
import win32con
import win32file
import win32process
from PySide6.QtCore import QThread, Signal

log = logging.getLogger(__name__)

# 按优先级排列：(COM ProgID, 对应进程名, 显示名)
CONVERTERS = [
    ("Word.Application", "winword.exe", "Microsoft Word"),
    ("KWPS.Application", "wps.exe", "WPS 文字"),
]

CONVERT_TIMEOUT_S = 60          # 单个文件最长转换时间
IDLE_QUIT_S = 60                # 空闲多久后关掉后台 Word
CACHE_MAX_AGE_DAYS = 7
CACHE_MAX_BYTES = 1 << 30       # 1 GB

PRIORITY_PRINT = 0              # 打印时的转换请求插队
PRIORITY_BACKGROUND = 1

# Word 常量
_wdExportFormatPDF = 17
_wdExportDocumentContent = 0     # 不带修订 / 批注标记
_wdExportDocumentWithMarkup = 7  # 带修订 / 批注标记
_wdDoNotSaveChanges = 0
_wdAlertsNone = 0
_msoAutomationSecurityForceDisable = 3
_wdRevisionsMarkupAll = 2
_wdRevisionsViewFinal = 0
_DRIVE_REMOTE = 4
_FAKE_PASSWORD = "__batchprint_no_password__"

NO_CONVERTER_MSG = "需要安装 Microsoft Word 或 WPS 才能打印 Word 文档"
ENCRYPTED_MSG = "文档已加密（设有打开密码或权限保护），无法自动转换"
CORRUPT_MSG = "文件已损坏，或不是有效的 Word 文档"

_ZIP_MAGIC = b"PK\x03\x04"
_CFB_MAGIC = bytes.fromhex("D0CF11E0A1B11AE1")   # OLE 复合文档：.doc，或加密后的 .docx


class ConversionError(Exception):
    """转换失败；str(exc) 是给用户看的说明。"""


class ConversionCancelled(Exception):
    """等待转换时被用户中止。"""


class _PasswordError(Exception):
    pass


class _TimeoutError(Exception):
    pass


# ---------------- 环境检测 ----------------
def available_converters() -> list[tuple[str, str, str]]:
    """本机已注册的转换器（按优先级）。"""
    found = []
    for progid, exe, name in CONVERTERS:
        try:
            winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + r"\CLSID"))
            found.append((progid, exe, name))
        except OSError:
            pass
    return found


def is_remote(path: Path) -> bool:
    """是否在网络位置（UNC 路径或映射的网络驱动器）。"""
    s = str(path)
    if s.startswith("\\\\"):
        return True
    drive = os.path.splitdrive(s)[0]
    if not drive:
        return False
    try:
        return win32file.GetDriveType(drive + "\\") == _DRIVE_REMOTE
    except Exception:  # noqa: BLE001
        return False


def check_word_file(src: Path) -> None:
    """转换前检查文件头，不合格直接抛 ConversionError。

    不能把格式不明的文件交给 Word：它会把文件当纯文本硬读，损坏文件会被
    "转换"成几十页乱码。空文件则会被转成一页白纸。
    """
    try:
        size = src.stat().st_size
        with open(src, "rb") as f:
            head = f.read(512)
    except OSError as exc:
        raise ConversionError(f"无法读取文件：{exc.strerror or exc}") from None
    if size == 0:
        raise ConversionError("文件是空的")
    if head.startswith(_CFB_MAGIC):
        if src.suffix.lower() == ".docx":      # 正常 .docx 是 ZIP；OLE 容器说明被加密了
            raise ConversionError(ENCRYPTED_MSG)
        return                                 # 老格式 .doc
    if head.startswith(_ZIP_MAGIC):
        try:
            with zipfile.ZipFile(src) as z:
                ok = "[Content_Types].xml" in z.namelist()
        except zipfile.BadZipFile:
            ok = False
        if not ok:
            raise ConversionError(CORRUPT_MSG)
        return
    text = head.lstrip(b"\xef\xbb\xbf").lstrip().lower()
    if text.startswith((b"{\\rtf", b"<", b"mime-version")):
        return                                 # 另存为 .doc 的 RTF / 网页格式，Word 能正常打开
    raise ConversionError(CORRUPT_MSG)


# ---------------- 缓存 ----------------
def work_dir() -> Path:
    """和 Word 交换文件（本地副本、导出的 PDF）用的临时目录。

    必须放在 %TEMP%，不能放在 %LOCALAPPDATA%：用微软商店版 Python 运行时，
    本程序写进 AppData 的文件会被系统悄悄重定向到商店应用的私有目录，
    Python 自己看得到，Word 却看不到，会报「找不到此文件」「目录名无效」。
    缓存目录只有本程序自己读写，不受这个影响。
    """
    d = Path(tempfile.gettempdir()) / "BatchPrint"
    d.mkdir(parents=True, exist_ok=True)
    return d


def cache_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home())
    d = Path(base) / "BatchPrint" / "cache" / "word"
    d.mkdir(parents=True, exist_ok=True)
    return d


def cache_path(src: Path, include_markup: bool) -> Path:
    st = src.stat()
    key = f"{str(src.resolve()).lower()}|{st.st_mtime_ns}|{st.st_size}|{int(include_markup)}"
    return cache_dir() / (hashlib.sha1(key.encode("utf-8")).hexdigest() + ".pdf")


def _pdf_pages(pdf: Path) -> int:
    with fitz.open(pdf) as doc:
        return doc.page_count


def lookup_cache(src: Path, include_markup: bool) -> Optional[tuple[Path, int]]:
    """缓存命中返回 (PDF 路径, 页数)，否则 None。"""
    try:
        pdf = cache_path(src, include_markup)
    except OSError:
        return None
    if not pdf.exists():
        return None
    try:
        pages = _pdf_pages(pdf)
    except Exception:  # noqa: BLE001 - 缓存文件损坏，删掉重转
        pdf.unlink(missing_ok=True)
        return None
    try:
        os.utime(pdf, None)          # 记录"最近使用"，供清理判断
    except OSError:
        pass
    return pdf, pages


def cleanup_cache() -> None:
    """删除 7 天未用的缓存；总量超过上限时按最久未用清理；清掉上次残留的临时文件。"""
    d = cache_dir()
    shutil.rmtree(work_dir(), ignore_errors=True)
    shutil.rmtree(d / "tmp", ignore_errors=True)      # 旧版本放在这里的临时目录
    now = time.time()
    files = []
    for p in d.glob("*.pdf"):
        try:
            st = p.stat()
        except OSError:
            continue
        if now - st.st_mtime > CACHE_MAX_AGE_DAYS * 86400:
            p.unlink(missing_ok=True)
        else:
            files.append((st.st_mtime, st.st_size, p))
    total = sum(size for _, size, _ in files)
    for _, size, p in sorted(files):
        if total <= CACHE_MAX_BYTES:
            break
        p.unlink(missing_ok=True)
        total -= size


# ---------------- 进程工具 ----------------
def _all_pids() -> set[int]:
    return set(win32process.EnumProcesses())


def _pid_exe(pid: int) -> str:
    try:
        h = win32api.OpenProcess(
            win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
    except Exception:  # noqa: BLE001
        return ""
    try:
        return os.path.basename(win32process.GetModuleFileNameEx(h, 0)).lower()
    except Exception:  # noqa: BLE001
        return ""
    finally:
        win32api.CloseHandle(h)


def _err_text(exc: BaseException) -> str:
    if isinstance(exc, pywintypes.com_error):
        info = exc.excepinfo
        if info and len(info) > 2 and info[2]:
            return str(info[2]).strip()
        return str(exc.strerror)
    return str(exc)


def _is_password_error(exc: BaseException) -> bool:
    text = _err_text(exc).lower()
    return "密码" in text or "password" in text


# ---------------- 一个后台 Word / WPS 实例 ----------------
class _Session:
    def __init__(self, progid: str, exe: str, name: str):
        self.progid, self.exe, self.name = progid, exe, name
        self.app = None
        self.pids: list[int] = []       # 本会话启动的进程（WPS 可能不止一个）
        self.killed = False

    def start(self) -> None:
        before = _all_pids()
        app = win32com.client.DispatchEx(self.progid)
        new = [p for p in _all_pids() - before if _pid_exe(p) == self.exe]
        if not new:
            # 没有新进程 = 接管了用户正在用的实例（WPS 可能这样）。
            # 不能隐藏或关闭用户的窗口，直接放弃这个转换器。
            del app
            raise ConversionError(f"{self.name} 正在使用中，无法在后台转换")
        self.pids = new
        self.app = app
        app.Visible = False
        for attr, value in (("DisplayAlerts", _wdAlertsNone),
                            ("AutomationSecurity", _msoAutomationSecurityForceDisable)):
            try:
                setattr(app, attr, value)
            except Exception:  # noqa: BLE001
                log.warning("%s 不支持设置 %s", self.name, attr)
        log.info("已启动后台 %s（PID %s）", self.name, self.pids)

    def alive(self) -> bool:
        try:
            _ = self.app.Name
            return True
        except Exception:  # noqa: BLE001
            return False

    def has_foreign_docs(self) -> bool:
        """后台实例里有没有不是我们打开的文档。

        后台 Word 运行期间，用户双击打开的文档有可能被系统塞进这个隐藏实例。
        """
        try:
            return int(self.app.Documents.Count) > 0
        except Exception:  # noqa: BLE001
            return False

    def hand_over(self) -> None:
        """把实例交还给用户：显示窗口，不关闭、不强杀，以免丢掉用户的文档。"""
        log.warning("后台 %s 里出现了用户的文档，已显示窗口交还给用户", self.name)
        try:
            self.app.Visible = True
        except Exception:  # noqa: BLE001
            pass
        self.app = None
        self.pids = []

    def convert(self, src: Path, dst: Path, include_markup: bool) -> None:
        watchdog = threading.Timer(CONVERT_TIMEOUT_S, self.kill)
        watchdog.daemon = True
        watchdog.start()
        doc = None
        try:
            # Open(文件名, 确认转换, 只读, 加入最近文件, 打开密码)
            doc = self.app.Documents.Open(str(src), False, True, False, _FAKE_PASSWORD)
            if include_markup:
                self._show_all_markup(doc)
            else:
                self._strip_markup(doc)
            item = _wdExportDocumentWithMarkup if include_markup else _wdExportDocumentContent
            doc.ExportAsFixedFormat(str(dst), _wdExportFormatPDF, False, 0, 0, 1, 1,
                                    item, True, True, 0, True, True, False)
        finally:
            if doc is not None:
                try:
                    doc.Close(_wdDoNotSaveChanges)
                except Exception:  # noqa: BLE001
                    pass
            watchdog.cancel()

    def _show_all_markup(self, doc) -> None:
        """导出带修订版本前，把视图设成"所有标记"，保证修订和批注都导出。"""
        try:
            view = doc.ActiveWindow.View
            view.RevisionsFilter.Markup = _wdRevisionsMarkupAll
            view.RevisionsFilter.View = _wdRevisionsViewFinal
            view.ShowRevisionsAndComments = True
        except Exception as exc:  # noqa: BLE001
            log.warning("设置修订视图失败: %s", exc)

    def _strip_markup(self, doc) -> None:
        """导出不带修订版本前：在内存里接受全部修订、删除批注。

        Word 的 wdExportDocumentContent 本身就不带标记，但 WPS 会忽略这个参数，
        照样把被删的文字印出来，所以统一在内存里处理一遍。文档是只读打开的，
        关闭时也不保存，原文件不受影响。
        """
        for action in (lambda: doc.Revisions.AcceptAll(),
                       lambda: doc.DeleteAllComments()):
            try:
                action()
            except Exception as exc:  # noqa: BLE001 - 如文档受编辑保护，退回只靠导出参数
                log.info("处理修订/批注失败（%s）: %s", self.name, _err_text(exc))

    def kill(self) -> None:
        """强制结束本会话启动的进程（超时保护用）。

        只结束启动时记下的 PID，并再次核对进程名，防止 PID 已被系统复用。
        """
        for pid in self.pids:
            if _pid_exe(pid) != self.exe:
                continue
            try:
                h = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
                win32api.TerminateProcess(h, 1)
                win32api.CloseHandle(h)
                self.killed = True
                log.warning("后台 %s 超时无响应，已结束进程 %s", self.name, pid)
            except Exception:  # noqa: BLE001
                log.exception("结束进程失败 %s", pid)

    def quit(self) -> None:
        if self.app is not None:
            if self.has_foreign_docs():
                self.hand_over()
                return
            try:
                self.app.Quit(_wdDoNotSaveChanges)
            except Exception:  # noqa: BLE001
                try:
                    self.app.Quit()
                except Exception:  # noqa: BLE001
                    pass
            self.app = None
        for _ in range(30):              # 最多等 3 秒正常退出
            if not set(self.pids) & _all_pids():
                return
            time.sleep(0.1)
        self.kill()


# ---------------- 请求 ----------------
class _Request:
    def __init__(self, key: str, src: Path, include_markup: bool):
        self.key = key
        self.src = src
        self.include_markup = include_markup
        self.event = threading.Event()
        self.result: Optional[tuple[Path, int]] = None
        self.error: Optional[str] = None
        self.cancelled = False


# ---------------- 转换线程 ----------------
class WordConverter(QThread):
    """唯一的转换线程。

    信号（在主线程接收）：
      converted(key, pdf 路径, 页数, 是否带修订)
      failed(key, 原因)
    """

    converted = Signal(str, str, int, bool)
    failed = Signal(str, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._converters = available_converters()
        self._q: queue.PriorityQueue = queue.PriorityQueue()
        self._seq = itertools.count()
        self._session: Optional[_Session] = None

    @property
    def available(self) -> bool:
        return bool(self._converters)

    # ---- 对外接口 ----
    def request(self, key: str, src: Path, include_markup: bool = False) -> None:
        """后台转换（导入后调用）；结果通过信号返回。"""
        self._q.put((PRIORITY_BACKGROUND, next(self._seq),
                     _Request(key, Path(src), include_markup)))

    def convert_blocking(self, key: str, src: Path, include_markup: bool,
                         cancel: Optional[Callable[[], bool]] = None) -> tuple[Path, int]:
        """打印线程调用：插队转换并等待结果；失败抛 ConversionError。"""
        hit = lookup_cache(Path(src), include_markup)
        if hit:
            return hit
        req = _Request(key, Path(src), include_markup)
        self._q.put((PRIORITY_PRINT, next(self._seq), req))
        while not req.event.wait(0.2):
            if cancel and cancel():
                req.cancelled = True
                raise ConversionCancelled()
        if req.error:
            raise ConversionError(req.error)
        return req.result

    def cancel_pending(self) -> None:
        """丢弃还没开始的后台请求（清空列表时用）。"""
        kept = []
        while True:
            try:
                item = self._q.get_nowait()
            except queue.Empty:
                break
            prio, seq, req = item
            if req is not None and prio == PRIORITY_BACKGROUND:
                req.cancelled = True
                req.event.set()
            else:
                kept.append(item)
        for item in kept:
            self._q.put(item)

    def shutdown(self, wait_ms: int = 8000) -> None:
        self.cancel_pending()
        self._q.put((-1, next(self._seq), None))
        self.wait(wait_ms)

    # ---- 线程主体 ----
    def run(self) -> None:  # noqa: D401 - QThread 入口
        pythoncom.CoInitialize()
        try:
            try:
                cleanup_cache()
            except Exception:  # noqa: BLE001
                log.exception("清理转换缓存失败")
            while True:
                try:
                    _, _, req = self._q.get(timeout=IDLE_QUIT_S)
                except queue.Empty:
                    self._close_session()          # 空闲了，关掉后台 Word 释放内存
                    continue
                if req is None:
                    break
                if req.cancelled:
                    req.event.set()
                    continue
                self._handle(req)
        finally:
            self._close_session()
            pythoncom.CoUninitialize()

    def _handle(self, req: _Request) -> None:
        try:
            req.result = self._convert(req.src, req.include_markup)
        except ConversionError as exc:
            req.error = str(exc)
        except Exception as exc:  # noqa: BLE001
            log.exception("转换出错: %s", req.src)
            req.error = f"转换失败：{_err_text(exc)}"
        req.event.set()
        if req.error:
            log.warning("转换失败 %s: %s", req.src, req.error)
            self.failed.emit(req.key, req.error)
        else:
            pdf, pages = req.result
            self.converted.emit(req.key, str(pdf), pages, req.include_markup)

    def _convert(self, src: Path, include_markup: bool) -> tuple[Path, int]:
        if not self._converters:
            raise ConversionError(NO_CONVERTER_MSG)
        if not src.exists():
            raise ConversionError("文件不存在")
        hit = lookup_cache(src, include_markup)
        if hit:
            return hit
        check_word_file(src)

        dst = cache_path(src, include_markup)
        tmp_dir = work_dir() / uuid.uuid4().hex
        tmp_dir.mkdir(parents=True)
        out = tmp_dir / "out.pdf"
        local_copy: Optional[Path] = None

        def get_copy() -> Path:
            nonlocal local_copy
            if local_copy is None:
                local_copy = tmp_dir / ("source" + src.suffix.lower())
                shutil.copyfile(src, local_copy)
            return local_copy

        try:
            errors = []
            # 网络盘文件：只用本地副本；本地文件：先直接转，打不开再用副本重试
            attempts = (True,) if is_remote(src) else (False, True)
            for progid, exe, name in self._converters:
                for use_copy in attempts:
                    try:
                        work = get_copy() if use_copy else src
                        t0 = time.perf_counter()
                        self._convert_with(progid, exe, name, work, out, include_markup)
                        # 临时目录和缓存目录可能不在同一"视角"下（见 work_dir），
                        # 先复制成 .part 再在缓存目录内原子改名，避免留下半个文件
                        part = dst.with_suffix(".part")
                        shutil.copyfile(out, part)
                        os.replace(part, dst)
                        log.info("已转换 %s（%s，%.1fs）", src.name, name,
                                 time.perf_counter() - t0)
                        return dst, _pdf_pages(dst)
                    except _PasswordError:
                        raise ConversionError(ENCRYPTED_MSG) from None
                    except _TimeoutError:
                        errors.append(f"{name} 超过 {CONVERT_TIMEOUT_S} 秒无响应")
                        break                       # 超时不再用副本重试，换下一个转换器
                    except ConversionError as exc:  # 如：WPS 正在使用中
                        errors.append(str(exc))
                        break
                    except Exception as exc:  # noqa: BLE001
                        errors.append(f"{name}：{_err_text(exc)}")
            unique = list(dict.fromkeys(errors))        # 去重，保留顺序
            raise ConversionError("转换失败（" + "；".join(unique) + "）")
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    def _convert_with(self, progid, exe, name, src: Path, dst: Path,
                      include_markup: bool) -> None:
        session = self._ensure_session(progid, exe, name)
        try:
            session.convert(src, dst, include_markup)
            if session.has_foreign_docs():      # 用户的文档跑进来了：交还，下次另起实例
                session.hand_over()
                self._session = None
        except Exception as exc:
            if session.killed:
                self._session = None
                raise _TimeoutError() from None
            if _is_password_error(exc):
                raise _PasswordError() from None
            if not session.alive():             # Word 崩溃 / 断开：丢弃，下次重启
                self._session = None
            raise

    def _ensure_session(self, progid, exe, name) -> _Session:
        if self._session and self._session.progid == progid and self._session.alive():
            return self._session
        self._close_session()
        session = _Session(progid, exe, name)
        try:
            session.start()
        except ConversionError:
            raise
        except Exception as exc:  # noqa: BLE001 - 启动失败与具体文档无关，直接换转换器
            raise ConversionError(f"无法启动 {name}：{_err_text(exc)}") from None
        self._session = session
        return session

    def _close_session(self) -> None:
        if self._session is not None:
            self._session.quit()
            log.info("已关闭后台 %s", self._session.name)
            self._session = None
