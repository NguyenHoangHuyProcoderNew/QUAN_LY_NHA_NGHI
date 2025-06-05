# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['APP_KHAI_BAO_LUU_TRU_2.py'],
    pathex=[],
    binaries=[],
    datas=[('logo_app.ico', '.'), ('done.wav', '.'), ('error.wav', '.'), ('loading.gif', '.')],
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
    [],
    exclude_binaries=True,
    name='Quan_Ly_Nha_Nghi',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['logo_app.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Quan_Ly_Nha_Nghi',
)
