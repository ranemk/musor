# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['l2_damage_meter.py'],
    pathex=[],
    binaries=[],
    datas=[('assets', 'assets'), ('win_ocr.ps1', '.')],
    hiddenimports=['tkinter', '_tkinter', 'tkinter.messagebox'],
    hookspath=['hooks'],
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
    name='parashaoly',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon='assets/parashaoly_circlet_x.ico',
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
