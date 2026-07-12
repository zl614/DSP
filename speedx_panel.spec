# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 打包配置 —— 由 GitHub Actions 在 windows-latest 上跑。

页面文件：仓库根目录下所有 *.html 自动打进 exe（新增页面不用改这个文件）。
运行时优先读 exe 同目录的 html，没有才用打进去的这份 —— 所以改页面不必重新打包。
"""
import os

# 自动收集根目录所有 html，避免哪天新增页面忘了加、或某个文件缺失导致构建失败
datas = [(f, ".") for f in os.listdir(".") if f.lower().endswith(".html")]
print("[spec] 打包这些页面:", [d[0] for d in datas])

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
    console=True,          # 保留黑框：启动时会打印地址、页面来源、报错信息，出问题好查
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
