# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

PROJECT_DIR = Path(SPEC).resolve().parent

# Files/folders required by the Flask application.
datas = [
    (str(PROJECT_DIR / "templates"), "templates"),
    (str(PROJECT_DIR / "static"), "static"),
    (str(PROJECT_DIR / "templates_docx"), "templates_docx"),
    (str(PROJECT_DIR / "database"), "database"),
]

# Add dynamically imported packages here only if PyInstaller reports them
# as missing during the build/test.
hiddenimports = [
    "flask",
    "jinja2",
    "werkzeug",
]

a = Analysis(
    [str(PROJECT_DIR / "app.py")],
    pathex=[str(PROJECT_DIR)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="EvidenceAutomation",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="EvidenceAutomation",
)
