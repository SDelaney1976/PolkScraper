# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['src/app.py'],
    pathex=['src'],
    binaries=[],
    datas=[('src/validator', 'validator'), ('src/capture', 'capture'), ('src/exports', 'exports')],
    hiddenimports=['capture', 'capture.gather_cases', 'capture.gather_case_details', 'exports', 'exports.export_data', 'exports.generate_english_letters', 'exports.generate_spanish_letters'],
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
    name='POLK County Scraper',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='POLK County Scraper',
)
app = BUNDLE(
    coll,
    name='POLK County Scraper.app',
    icon=None,
    bundle_identifier=None,
)
