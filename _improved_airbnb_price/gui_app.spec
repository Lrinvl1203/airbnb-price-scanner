# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for Airbnb 가격 분석 GUI
"""

import sys
from pathlib import Path
import babel
import tkcalendar

BABEL_DATA = str(Path(babel.__file__).parent / "global.dat")
BABEL_DIR  = str(Path(babel.__file__).parent)
TKCAL_DIR  = str(Path(tkcalendar.__file__).parent)

block_cipher = None

a = Analysis(
    ["gui_app.py"],
    pathex=["."],
    binaries=[],
    datas=[
        (BABEL_DIR,  "babel"),
        (TKCAL_DIR,  "tkcalendar"),
    ],
    hiddenimports=[
        # curl_cffi
        "curl_cffi",
        "curl_cffi.requests",
        "curl_cffi._wrapper",
        "_cffi_backend",
        # tkcalendar
        "tkcalendar",
        "babel",
        "babel.numbers",
        "babel.dates",
        # xlsxwriter
        "xlsxwriter",
        # lxml (airbnb_fetch HTML parsing)
        "lxml",
        "lxml.etree",
        "lxml.html",
        # tkinter
        "tkinter",
        "tkinter.ttk",
        "tkinter.messagebox",
        "tkinter.filedialog",
        # stdlib extras
        "queue",
        "threading",
        "importlib",
        "statistics",
        "zipfile",
        # local modules
        "airbnb_fetch",
        "market_report",
        "export_excel",
        "export_excel_detail",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["flask", "matplotlib", "numpy", "pandas", "scipy", "PIL.ImageTk"],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="에어비앤비_시장분석",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,           # GUI 앱 → 콘솔창 없음
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="에어비앤비_시장분석",
)
