"""资源路径：兼容开发运行与 PyInstaller 打包后（onefile / onedir）。"""
from __future__ import annotations

import os
import sys


def resource_path(rel: str) -> str:
    """返回资源的绝对路径。

    打包后 PyInstaller 会设置 sys._MEIPASS（onefile 为临时目录，onedir 为 _internal）；
    开发运行时以项目根目录为基准。
    """
    base = getattr(sys, "_MEIPASS", None)
    if base is None:
        # 项目根 = 本文件所在包(batchprint)的上级目录
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def app_icon_path() -> str:
    return resource_path(os.path.join("assets", "icon.ico"))
