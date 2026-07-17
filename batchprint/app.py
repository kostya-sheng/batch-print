"""程序入口。"""
from __future__ import annotations

import sys

import os

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from .logging_setup import setup_logging
from .resources import app_icon_path
from .ui.main_window import MainWindow


def main() -> int:
    setup_logging()
    app = QApplication(sys.argv)
    app.setApplicationName("批量打印")
    icon_file = app_icon_path()
    if os.path.exists(icon_file):
        app.setWindowIcon(QIcon(icon_file))
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
