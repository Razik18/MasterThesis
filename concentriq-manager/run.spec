# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import collect_all

datas = [('C:\\Users\\ME4A4TQ\\Documents\\pyth\\concentriq_manager\\concentriq-manager\\ConcentriqSDK', 'ConcentriqSDK/'), ('C:\\Users\\ME4A4TQ\\Documents\\pyth\\concentriq_manager\\concentriq-manager\\app', 'app/')]
binaries = []
hiddenimports = ['win32timezone', 'imghdr', 'imgaug', 'pyclipper']
datas += collect_data_files('paddle')
tmp_ret = collect_all('paddleocr')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['C:\\Users\\ME4A4TQ\\Documents\\pyth\\concentriq_manager\\concentriq-manager\\run.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='run',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['C:\\Users\\ME4A4TQ\\Downloads\\logo.ico'],
)
