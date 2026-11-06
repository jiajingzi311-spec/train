# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['MRDN物料转置.py'],
    pathex=[],
    binaries=[('D:\\Program Files\\GDEClaw\\Library\\bin\\ffi.dll', '.'), ('D:\\Program Files\\GDEClaw\\Library\\bin\\libexpat.dll', '.'), ('D:\\Program Files\\GDEClaw\\Library\\bin\\libcrypto-3-x64.dll', '.'), ('D:\\Program Files\\GDEClaw\\Library\\bin\\libssl-3-x64.dll', '.'), ('D:\\Program Files\\GDEClaw\\Library\\bin\\liblzma.dll', '.'), ('D:\\Program Files\\GDEClaw\\Library\\bin\\libbz2.dll', '.')],
    datas=[],
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
    name='mrdn_transfor',
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
)
