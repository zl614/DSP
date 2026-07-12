# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec -- built by GitHub Actions on windows-latest.
#
# NOTE: keep this file pure ASCII. On Windows CI, stdout is a pipe and Python
# falls back to cp1252, so printing non-ASCII here raises UnicodeEncodeError
# and PyInstaller dies before it even starts building.
#
# All *.html in the repo root are collected automatically, so adding a new
# page needs no change here. At runtime the app prefers an .html sitting next
# to the exe, and only falls back to the bundled copy -- so editing a page
# does not require a rebuild.

import os

datas = [(f, ".") for f in os.listdir(".") if f.lower().endswith(".html")]
print("[spec] bundling pages:", [d[0] for d in datas])

a = Analysis(
    ["speedx_relay.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=["openpyxl"],
    hookspath=[],
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "numpy", "pandas", "PIL"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="SpeedX-Panel",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
