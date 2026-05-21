# -*- mode: python ; coding: utf-8 -*-
# Busca Branch Publicada — build v4 (PyInstaller)

a = Analysis(
    ['index.py'],
    pathex=[],
    binaries=[],
    datas=[('logo.jpeg', '.')],
    hiddenimports=['openpyxl', 'PIL', 'PIL.Image', 'PIL.ImageTk'],
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
    name='BuscaBranchGit_v4',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
