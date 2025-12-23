# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['file_converter_gui_v3.0 (UI improvements).py'],
    pathex=[],
    binaries=[],
    datas=[('assets\\\\logo100.png', 'assets'), ('assets\\\\drag100.png', 'assets'), ('assets\\\\app.ico', 'assets'), ('assets\\\\GRF_theme.json', 'assets'), ('assets\\\\busy_splash.png', 'assets')],
    hiddenimports=[],
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
    name='File_Converter_v3.0',
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
    icon=['assets\\app.ico'],
)
