# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for trajectory_splitter web GUI (v1.0).

Build (from trajectory_splitter/):
  pyinstaller trajectory_splitter_web.spec --noconfirm --clean

Output:
  dist/trajectory_splitter_web.exe  (onefile, single distributable)
"""

import os
import sys

block_cipher = None

# spec 文件所在目录 = trajectory_splitter/
SPEC_DIR = os.path.dirname(os.path.abspath(SPEC))


datas = [
    # web/ 前端 (HTML/CSS/JS) 跟着打进 exe, 运行时 web_gui.py 走 sys._MEIPASS 找
    (os.path.join(SPEC_DIR, 'web'), 'web'),
]

# pywebview 6.x 在 Windows 走 edge webview2 + clr_loader/pythonnet (走 .NET)
# pyinstaller-hooks-contrib 自带 webview 的 hook, 但 hiddenimports 写一份保险
hiddenimports = [
    'webview',
    'webview.platforms.edge',
    'webview.platforms.winforms',
    'clr_loader',
    'pythonnet',
    # xls_io 走 Excel COM 读写 .xls
    'win32com',
    'win32com.client',
    'pythoncom',
    'pywintypes',
]

a = Analysis(
    [os.path.join(SPEC_DIR, 'web_gui.py')],
    pathex=[SPEC_DIR],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 不需要的: 老 GUI + 测试 + matplotlib 之类全剔除
        'matplotlib',
        'numpy',
        'pandas',
        'PIL',
        'PySide6',
        'tkinter.test',
        'unittest',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='trajectory_splitter_web',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # windowed app, 不弹黑窗
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)