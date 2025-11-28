# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
    ('version 2.1.0.txt', '.'),
    ('config.yaml', '.'),
    ('updater.exe', '.'),
    ('templates', 'templates'),
    ],
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
    name='Stockwise',
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
    icon=['resources/icon.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Stockwise',
)

# ----------------- Post-build copying -----------------

import shutil
import os

dist_dir = os.path.join('dist', 'Stockwise')
internal_dir = os.path.join(dist_dir, '_internal')

files_to_copy = [
    'config.yaml',
    'version 2.1.0.txt',
    'updater.exe',
    'templates'
]

for filename in files_to_copy:
    src = os.path.join(internal_dir, filename)
    dst = os.path.join(dist_dir, filename)
    try:
        if os.path.exists(src):
            if os.path.isdir(src):
                if os.path.exists(dst):
                    shutil.rmtree(dst)
                shutil.copytree(src, dst)
            else:
                shutil.copy2(src, dst)

            print(f"[Post-Build] {filename} copied to {dist_dir}")
        else:
            print(f"[Post-Build] {filename} not found in _internal")

    except Exception as e:
        print(f"[Post-Build] Error copying {filename}: {e}")