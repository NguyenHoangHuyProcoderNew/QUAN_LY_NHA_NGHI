# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['APP_KHAI_BAO_LUU_TRU_2.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('logo_app.ico', '.'),
        ('done.wav', '.'),
        ('error.wav', '.'),
        ('config_cam.json', '.'),
    ],
    hiddenimports=['win32timezone'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'pandas', 'PIL', 'numpy', 'tkinter'],
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
    name='APP_KHAI_BAO_LUU_TRU_2',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='logo_app.ico',
    version='file_version_info.txt',
)