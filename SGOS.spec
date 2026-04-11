# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = ['sgos_web.app', 'sgos_web.engine', 'sgos_web.extensions', 'sgos_web.comps_engine', 'sgos_web.comps_routes']
hiddenimports += collect_submodules('clr_loader')
hiddenimports += collect_submodules('pythonnet')


a = Analysis(
    ['desktop.py'],
    pathex=[],
    binaries=[],
    datas=[('sgos_web/templates', 'sgos_web/templates'), ('sgos_web/static', 'sgos_web/static'), ('.env', '.')],
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
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SGOS',
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
    name='SGOS',
)
