# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller build spec for Super Filter.
#
# Build:   pyinstaller app.spec --noconfirm
# Output:  dist/SuperFilter/        (Windows/Linux: a folder you can zip and copy)
#          dist/Super Filter.app    (macOS: a double-clickable app bundle)
#
# WHY onedir and not onefile: onefile unpacks ~70 MB to a temp directory on every
# launch, which adds several seconds of apparent hang before the window appears.
# A folder starts immediately and still copies as one unit.
#
# WHY datas matters: report_template.xlsx is NOT code, so PyInstaller cannot find it
# by following imports. Without this line the app builds cleanly and then fails at
# export with FileNotFoundError. That is exactly what broke on 2026-08-05.

datas = [
    ('report_template.xlsx', '.'),
]

hiddenimports = [
    'super_filter',   # imported by app.py; listed explicitly so a refactor cannot drop it
]

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['streamlit', 'altair', 'pyarrow', 'matplotlib', 'IPython', 'pytest'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SuperFilter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
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
    name='SuperFilter',
)

app = BUNDLE(
    coll,
    name='Super Filter.app',
    icon=None,                      # supply an .icns here once one exists
    bundle_identifier='com.ammar.superfilter',
    info_plist={
        'CFBundleShortVersionString': '1.0.0',
        'NSHighResolutionCapable': True,
    },
)
