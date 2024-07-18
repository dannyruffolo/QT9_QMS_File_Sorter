# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    [r'C:\Users\druffolo\Desktop\QT9_QMS_File_Sorter\main.py'],
    pathex=[],
    binaries=[],
    datas=[(r'C:\Users\druffolo\Desktop\QT9_QMS_File_Sorter\config.py', '.')],
    hiddenimports=['plyer.platforms.win.notification'],
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
    name='QT9 QMS File Sorter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[r'C:\Users\druffolo\Desktop\File Sorter Installer & EXE Files\app_icon.ico'],
)
