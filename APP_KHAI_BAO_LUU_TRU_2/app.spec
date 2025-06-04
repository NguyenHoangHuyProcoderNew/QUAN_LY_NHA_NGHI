# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['APP_KHAI_BAO_LUU_TRU_2.py'],
    pathex=[],
    binaries=[
        ('libiconv.dll', '.'),
        ('libzbar-64.dll', '.')
    ],
    datas=[
        ('logo_app.ico', '.'),
        ('done.wav', '.'),
        ('error.wav', '.'),
        ('config_cam.json', '.')
    ],
    hiddenimports=[
        'win32com.client',
        'win32timezone',
        'cv2',
        'numpy',
        'PyQt5',
        'PyQt5.QtMultimedia',
        'PyQt5.QtMultimediaWidgets',
        'qreader',
        'unidecode',
        'psutil',
        'pyzbar',
        'PIL',
        'scipy'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'tkinter',
        'pandas',
        'PyQt5.QtWebEngine',
        'PyQt5.QtWebEngineCore',
        'PyQt5.QtWebEngineWidgets'
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Loại bỏ các file không cần thiết để giảm kích thước
a.binaries = [x for x in a.binaries if not x[0].startswith('mfc')]
a.binaries = [x for x in a.binaries if not x[0].startswith('opengl')]
a.binaries = [x for x in a.binaries if not x[0].startswith('qt5web')]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='APP_KHAI_BAO_LUU_TRU_2',
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
    icon='logo_app.ico',
    uac_admin=True
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='APP_KHAI_BAO_LUU_TRU_2',
)