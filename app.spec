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

import sys

# All three icon files are rendered from icon.svg by make_icons.sh. PyInstaller wants
# a .ico to embed in a Windows .exe and a .icns for a macOS bundle; icon.png is the
# one Tk itself can load at runtime for the window icon, so it ships as data.
EXE_ICON = 'icon.ico' if sys.platform == 'win32' else 'icon.icns'

datas = [
    ('report_template.xlsx', '.'),
    ('icon.png', '.'),
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
    # pandas and numpy were dropped when the GUI stopped doing its own maths.
    # Excluding them explicitly keeps the build around 25 MB instead of 70 MB.
    excludes=['pandas', 'numpy', 'streamlit', 'altair', 'pyarrow',
              'matplotlib', 'IPython', 'pytest'],
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
    icon=EXE_ICON,
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
    icon='icon.icns',
    bundle_identifier='com.ammar.superfilter',
    info_plist={
        'CFBundleShortVersionString': '1.0.0',
        'NSHighResolutionCapable': True,
    },
)
