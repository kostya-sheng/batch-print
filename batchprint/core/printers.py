"""打印机管理：枚举打印机、获取默认打印机、调出原生设置对话框拿 DEVMODE。"""
from __future__ import annotations

import logging
from typing import Any, Optional

import win32print

log = logging.getLogger(__name__)

# DocumentProperties 的 fMode 标志（win32con 不一定全有，这里显式定义）
DM_OUT_BUFFER = 2   # 把结果写回输出 DEVMODE
DM_IN_PROMPT = 4    # 弹出打印机属性对话框
DM_IN_BUFFER = 8    # 以输入 DEVMODE 为初始值
IDOK = 1


def list_printers() -> list[str]:
    """返回本机可用打印机名称列表（本地 + 网络连接）。"""
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    try:
        printers = win32print.EnumPrinters(flags)
    except Exception as exc:  # noqa: BLE001
        log.error("枚举打印机失败: %s", exc)
        return []
    return [p[2] for p in printers]


def default_printer() -> Optional[str]:
    try:
        return win32print.GetDefaultPrinter()
    except Exception:  # noqa: BLE001
        return None


def get_default_devmode(printer_name: str) -> Optional[Any]:
    """取打印机当前默认 DEVMODE（可能为 None）。"""
    hprinter = win32print.OpenPrinter(printer_name)
    try:
        info = win32print.GetPrinter(hprinter, 2)
        return info.get("pDevMode")
    except Exception as exc:  # noqa: BLE001
        log.warning("获取默认 DEVMODE 失败 %s: %s", printer_name, exc)
        return None
    finally:
        win32print.ClosePrinter(hprinter)


def clone_devmode(printer_name: str, source: Optional[Any]) -> Optional[Any]:
    """返回 source 的**独立副本**（不共享同一对象）。

    DocumentProperties 会就地修改 DEVMODE，若全局与单文档共用同一对象会相互串改；
    因此在弹对话框或保存前先克隆出独立缓冲。source 为 None 返回 None。
    """
    if source is None:
        return None
    hprinter = win32print.OpenPrinter(printer_name)
    try:
        work = get_default_devmode(printer_name)   # 全新的独立对象
        if work is None:
            return None
        # 把 source 的设置合并进独立缓冲（复制值，不共享对象）
        win32print.DocumentProperties(
            0, hprinter, printer_name,
            work,      # output（独立对象）
            source,    # input（要复制的设置）
            DM_IN_BUFFER | DM_OUT_BUFFER,
        )
        return work
    except Exception as exc:  # noqa: BLE001
        log.warning("克隆 DEVMODE 失败 %s: %s", printer_name, exc)
        return None
    finally:
        win32print.ClosePrinter(hprinter)


def open_printer_dialog(hwnd: int, printer_name: str,
                        devmode_in: Optional[Any] = None) -> Optional[Any]:
    """调出打印机原生属性对话框（双面/彩色/纸张等）。

    返回用户确认后的 DEVMODE（一个**独立对象**）；用户取消或失败返回 None。
    hwnd: 父窗口句柄（可用 0）。
    devmode_in: 作为初始值的 DEVMODE，None 则取打印机默认。
    """
    hprinter = win32print.OpenPrinter(printer_name)
    try:
        # 始终在独立缓冲上操作，避免就地改到调用方（全局/其它文档）的对象
        if devmode_in is not None:
            work = clone_devmode(printer_name, devmode_in)
        else:
            work = get_default_devmode(printer_name)
        if work is None:
            log.error("无法获得初始 DEVMODE，打印机 %s 可能不支持。", printer_name)
            return None

        ret = win32print.DocumentProperties(
            hwnd,
            hprinter,
            printer_name,
            work,   # output
            work,   # input
            DM_IN_PROMPT | DM_IN_BUFFER | DM_OUT_BUFFER,
        )
        if ret == IDOK:
            return work
        return None
    except Exception as exc:  # noqa: BLE001
        log.error("打开打印机属性对话框失败 %s: %s", printer_name, exc)
        return None
    finally:
        win32print.ClosePrinter(hprinter)
