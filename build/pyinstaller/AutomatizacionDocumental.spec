# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

project_root = Path.cwd().resolve()
src_root = project_root / "src"

datas = [
    (str(project_root / "config.example.json"), "."),
    (str(project_root / "README.md"), "."),
    (str(project_root / "assets" / "app_icon.png"), "assets"),
    (str(project_root / "assets" / "app_icon.ico"), "assets"),
]

hiddenimports = [
    "win32com",
    "win32com.client",
    "pythoncom",
    "docx2pdf",
]

a = Analysis(
    [str(project_root / "run_app.py")],
    pathex=[str(project_root), str(src_root)],
    binaries=[],
    datas=datas,
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
    name="AutomatizacionDocumental",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=str(project_root / "assets" / "app_icon.ico"),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="AutomatizacionDocumental",
)
