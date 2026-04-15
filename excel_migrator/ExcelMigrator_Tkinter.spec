# -*- mode: python ; coding: utf-8 -*-
"""
Excel 数据迁移工具 - PyInstaller 打包配置 (tkinter + ttkbootstrap)
"""
import os
import sys

block_cipher = None

current_dir = os.path.dirname(os.path.abspath(SPEC))

hiddenimports = [
    'ttkbootstrap',
    'openpyxl',
    'openpyxl.utils',
    'openpyxl.utils.cell',
    'openpyxl.workbook',
    'openpyxl.worksheet',
]

datas = [
    (os.path.join(current_dir, 'config.py'), '.'),
    (os.path.join(current_dir, 'launcher.py'), '.'),
    (os.path.join(current_dir, 'updater.py'), '.'),
    (os.path.join(current_dir, 'gui.py'), '.'),
    (os.path.join(current_dir, 'MANUAL.md'), '.'),
    (os.path.join(current_dir, 'VERSION.md'), '.'),
]

for subdir in ['core', 'storage', 'ui', 'utils']:
    src_dir = os.path.join(current_dir, subdir)
    if os.path.exists(src_dir):
        datas.append((src_dir, subdir))

a = Analysis(
    ['launcher.py'],
    pathex=[current_dir],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PyQt5', 'PyQt5.QtCore', 'PyQt5.QtGui', 'PyQt5.QtWidgets', 'PyQt5.sip'],
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
    a.zipfiles,
    a.datas,
    [],
    name='ExcelMigrator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    runtime_tmpdir=None,
)
