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
        ('loading.gif', '.'),
    ],
    hiddenimports=[
        'win32timezone',
        'cv2',
        'numpy',
        'PyQt5',
        'PyQt5.QtMultimedia',
        'PyQt5.QtMultimediaWidgets'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

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
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='logo_app.ico'
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Quan_Ly_Nha_Nghi',
) 