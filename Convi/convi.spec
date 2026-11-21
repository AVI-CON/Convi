# -*- mode: python ; coding: utf-8 -*-
# Облегчённая сборка: лишние части Qt и переводы не попадают в exe.
from PyInstaller.utils.hooks import collect_all
datas, binaries = [], []
hidden = ["PySide6.QtNetwork", "PySide6.QtSvg", "markdown", "send2trash", "win32com.client", "pythoncom"]
for pkg in ("imageio_ffmpeg", "pillow_heif", "docx", "pptx", "openpyxl", "reportlab", "pymupdf", "fitz"):
    try:
        d, b, h = collect_all(pkg); datas += d; binaries += b; hidden += h
    except Exception: pass
EXCL = ["tkinter", "test", "pydoc_data", "lib2to3", "setuptools", "pkg_resources", "PySide6.QtWebEngineCore", "PySide6.QtWebEngineWidgets",
        "PySide6.QtQuick", "PySide6.QtQml", "PySide6.Qt3DCore", "PySide6.QtMultimedia", "PySide6.QtCharts", "PySide6.QtPdf", "PySide6.QtDesigner"]
a = Analysis(["convi.py"], pathex=[], binaries=binaries, datas=datas, hiddenimports=hidden, hookspath=[], runtime_hooks=[], excludes=EXCL, noarchive=False)
DROP = ("opengl32sw", "qt6pdf", "qt6quick", "qt6qml", "qt6virtualkeyboard", "qt6designer", "qt6webengine")
a.binaries = [x for x in a.binaries if not any(s in x[0].lower() for s in DROP)]
a.datas = [x for x in a.datas if "translations" not in x[0].lower().replace("\\", "/")]
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], name="Convi", icon="convi.ico", debug=False, strip=False, upx=True,
          upx_exclude=["vcruntime140.dll", "vcruntime140_1.dll", "Qt6Core.dll", "Qt6Gui.dll", "Qt6Widgets.dll", "Qt6Network.dll", "Qt6Svg.dll"],
          runtime_tmpdir=None, console=False)
