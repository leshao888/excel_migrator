# -*- mode: python ; coding: utf-8 -*-
import os
import sys

block_cipher = None

# 获取当前目录
current_dir = os.path.dirname(os.path.abspath(SPEC))

# 收集所有模块
hiddenimports = [
    'streamlit',
    'streamlit.runtime',
    'streamlit.runtime.scriptrunner',
    'streamlit.runtime.legacy_caching',
    'streamlit.web',
    'streamlit.web.cli',
    'streamlit.web.server',
    'streamlit.web.server.ws_manager',
    'streamlit.elements',
    'streamlit.commands',
    'streamlit.components',
    'streamlit.components.v1',
    'openpyxl',
    'openpyxl.utils',
    'openpyxl.utils.cell',
    'openpyxl.workbook',
    'openpyxl.worksheet',
    'pandas',
    'pandas.core',
    'numpy',
    'numpy.core',
    'altair',
    'altair.vegalite',
    'tzdata',
]

# 添加 streamlit 所有子模块
try:
    import streamlit
    import pkgutil
    for importer, modname, ispkg in pkgutil.iter_modules(streamlit.__path__):
        if modname not in hiddenimports:
            hiddenimports.append(f'streamlit.{modname}')
except:
    pass

# 收集数据文件
datas = [
    (os.path.join(current_dir, 'config.py'), '.'),
    (os.path.join(current_dir, 'main.py'), '.'),
    (os.path.join(current_dir, 'MANUAL.md'), '.'),
    (os.path.join(current_dir, 'VERSION.md'), '.'),
]

# 添加所有子目录
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

# 打包成单个exe文件
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
    console=True,
    disable_windowed_traceback=False,
    runtime_tmpdir=None,
)
