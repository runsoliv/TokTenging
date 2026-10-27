# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


a = Analysis(
    ['TokTenging.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('temp2.png', 'tok'),
        ('success.wav', 'tok'),
        ('error.wav', 'tok'),
        ('Tok-Tenging.ico', 'tok'),
        *(
            [(str(Path('data folder') / 'data codex.xlsx'), 'tok\\data folder')]
            if (Path('data folder') / 'data codex.xlsx').exists()
            else []
        ),
    ],
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
    icon='Tok-Tenging.ico',
)
