# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：单文件、无控制台窗口，裁掉用不到的 Qt 大模块。"""

# 我们只用 QtWidgets/QtGui/QtCore，排除 PySide6 里体积大的无关模块以减小 exe。
_excludes = [
    "PySide6.QtWebEngineCore", "PySide6.QtWebEngineWidgets", "PySide6.QtWebEngineQuick",
    "PySide6.QtQml", "PySide6.QtQuick", "PySide6.QtQuick3D", "PySide6.QtQuickWidgets",
    "PySide6.QtQuickControls2",
    "PySide6.Qt3DCore", "PySide6.Qt3DRender", "PySide6.Qt3DInput", "PySide6.Qt3DLogic",
    "PySide6.Qt3DAnimation", "PySide6.Qt3DExtras",
    "PySide6.QtCharts", "PySide6.QtDataVisualization", "PySide6.QtGraphs",
    "PySide6.QtMultimedia", "PySide6.QtMultimediaWidgets", "PySide6.QtSpatialAudio",
    "PySide6.QtBluetooth", "PySide6.QtNfc", "PySide6.QtPositioning", "PySide6.QtLocation",
    "PySide6.QtSensors", "PySide6.QtSerialPort", "PySide6.QtSerialBus",
    "PySide6.QtWebSockets", "PySide6.QtWebChannel", "PySide6.QtWebView",
    "PySide6.QtHttpServer", "PySide6.QtNetworkAuth",
    "PySide6.QtSql", "PySide6.QtTest", "PySide6.QtDesigner", "PySide6.QtHelp",
    "PySide6.QtUiTools", "PySide6.QtScxml", "PySide6.QtStateMachine",
    "PySide6.QtRemoteObjects", "PySide6.QtTextToSpeech", "PySide6.QtPdf",
    "PySide6.QtPdfWidgets",
    "tkinter",
]

a = Analysis(
    ["run.py"],
    pathex=[],
    binaries=[],
    datas=[("assets/icon.ico", "assets")],
    hiddenimports=[
        "win32ui", "win32gui", "win32print", "win32con", "win32api",
        # Word 文档转换（COM 自动化）
        "win32com", "win32com.client", "pythoncom", "pywintypes",
        "win32file", "win32process",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=_excludes,
    noarchive=False,
)

pyz = PYZ(a.pure)

# 目录版（onedir / 绿色软件）：exe 不内嵌依赖，依赖由 COLLECT 收进同一文件夹。
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="批量打印",
    icon="assets/icon.ico",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="批量打印",
)
