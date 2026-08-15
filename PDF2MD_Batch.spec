# -*- mode: python ; coding: utf-8 -*-

import os
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files


project_root = Path(SPECPATH)
tesseract_dir = Path(
    os.environ.get("PDF2MD_TESSERACT_RUNTIME", project_root / "build" / "tesseract_runtime")
)
if not (tesseract_dir / "tessdata" / "spa.traineddata").is_file():
    raise SystemExit(
        "Falta build/tesseract_runtime. Ejecuta scripts/build_release.ps1 para preparar OCR."
    )

datas = [
    (str(project_root / "assets"), "assets"),
    (str(tesseract_dir), "tesseract"),
]
datas += collect_data_files("pymupdf", includes=["layout/resources/**"])

hiddenimports = ["pystray._win32"]

a = Analysis(
    ["launcher.py"],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "gunicorn",
        "torch",
        "torchvision",
        "torchaudio",
        "pandas",
        "pyarrow",
        "matplotlib",
        "scipy",
        "gi",
    ],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="PDF2MD",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    icon=str(project_root / "assets" / "pdf2md.ico"),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="PDF2MD",
)
