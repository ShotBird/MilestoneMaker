# -*- mode: python ; coding: utf-8 -*-
# Build:  pyinstaller --noconfirm milestone_maker.spec
# Bump APP_VER here AND VERSION in main.py together.
from PyInstaller.utils.hooks import collect_submodules

APP_VER = "v1.0"
APP_NAME = "Milestone Maker(%s)" % APP_VER

hiddenimports = ['lxml._elementpath']
hiddenimports += collect_submodules('pptx')
hiddenimports += collect_submodules('openpyxl')
hiddenimports += collect_submodules('comtypes')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('assets\\output.pptx', '.'), ('assets\\template_input.xlsm', '.'), ('assets\\app.ico', '.')],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                 # --windowed (GUI app, no console)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets\\app.ico',
)
