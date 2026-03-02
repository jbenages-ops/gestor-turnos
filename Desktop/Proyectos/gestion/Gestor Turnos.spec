# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['generador_turnos.py'],
    pathex=['/Users/juanbenages/Desktop/Proyectos/gestion'],
    binaries=[],
    datas=[],
    hiddenimports=[
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'pandas',
        'calendar',
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
    name='Gestor Turnos',
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
    icon='icon.icns',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Gestor Turnos',
)

app = BUNDLE(
    coll,
    name='Gestor Turnos.app',
    icon='icon.icns',
    bundle_identifier='com.juanbenages.gestorturnos',
    info_plist={
        'CFBundleName':              'Gestor Turnos',
        'CFBundleDisplayName':       'Gestor Turnos',
        'CFBundleShortVersionString': '1.0.0',
        'CFBundleVersion':           '1.0.0',
        'NSHighResolutionCapable':   True,
        'LSUIElement':               False,
    },
)
