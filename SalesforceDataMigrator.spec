# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['record_migrator\\main.py'],
    pathex=[],
    binaries=[],
    datas=[('record_migrator\\config.yaml', 'record_migrator'), ('record_migrator\\import_config.yaml', 'record_migrator')],
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
    name='SalesforceDataMigrator',
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
    version='version.txt',
    icon=['C:\\Users\\simon\\Pictures\\sfondi\\Foto Milano-8.jpg'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SalesforceDataMigrator',
)
