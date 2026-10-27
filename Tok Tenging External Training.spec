# -*- mode: python ; coding: utf-8 -*-

import sys
from pathlib import Path


python_base = Path(sys.base_prefix)
tcl_root = python_base / 'tcl'

binaries = []
for dll_name in ('tcl86t.dll', 'tk86t.dll'):
    dll_path = python_base / 'DLLs' / dll_name
    if dll_path.exists():
        binaries.append((str(dll_path), '.'))

datas = [
    ('tok\\temp2.png', 'tok'),
    ('tok\\success.wav', 'tok'),
    ('tok\\error.wav', 'tok'),
    ('tok\\Tok-Tenging.ico', 'tok'),
]

for source, target in (
    (python_base / 'Lib' / 'tkinter', 'py_lib\\tkinter'),
    (tcl_root / 'tcl8.6', '_tcl_data'),
    (tcl_root / 'tk8.6', '_tk_data'),
):
    if source.exists():
        datas.append((str(source), target))


a = Analysis(
    ['tok\\TokTenging.py'],
    pathex=['tok'],
    binaries=binaries,
    datas=datas,
    hiddenimports=[
        '_tkinter',
        'auto_coder',
        'bank_detection',
        'bank_formatter',
        'bank_utils',
        'coded_file_tools',
        'email_downloader',
        'folder_printer',
        'ml_artifacts',
        'ml_dataset',
        'ml_evaluation',
        'ml_features',
        'ml_model',
        'ml_runtime',
        'ml_training',
        'shareholder_registry',
        'shareholder_registry_data',
        'training_archive',
        'ui_components',
        'pypdf',
        'tkinterdnd2',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['pyinstaller_tcl_runtime.py'],
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
    name='Tok Tenging',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='tok\\Tok-Tenging.ico',
)
