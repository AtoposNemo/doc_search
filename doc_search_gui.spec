# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for Document Search GUI (tkinter, Win7-compatible, onefile)."""

block_cipher = None

a = Analysis(
    ['main_gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'search_docs',
        'win32com',
        'win32com.client',
        'pythoncom',
        'pywintypes',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter.test',
        'unittest',
        'pydoc',
        'doctest',
        'pdb',
        'profile',
        'pstats',
        'difflib',
        'flask',
        'werkzeug',
        'jinja2',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# Onefile mode: everything bundled into a single exe
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='文档搜索',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Windowed app: no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
