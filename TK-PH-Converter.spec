# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for the TK Philippines table converter.

Build:
    pyinstaller --noconfirm TK-PH-Converter.spec
or use build_windows.bat which wraps this and copies the assets/ folder
next to the produced exe.
"""

from pathlib import Path
import sys

block_cipher = None
APP_NAME = "TK-PH-Converter"
PROJECT_DIR = Path(SPECPATH).resolve()  # set by PyInstaller

# Entry point: the GUI main() in app/main.py
a = Analysis(
    [str(PROJECT_DIR / "app" / "main.py")],
    pathex=[str(PROJECT_DIR)],
    binaries=[],
    datas=[],  # assets/ is copied next to the exe by build_windows.bat
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Trim unused heavy modules to keep the exe lean
        "matplotlib", "scipy", "pandas", "numpy.testing",
        "pytest", "setuptools",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # GUI app: no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,               # add an .ico path here if you have one
)
