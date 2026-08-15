"""Shared runtime paths and metadata for source and packaged execution."""

from __future__ import annotations

import os
import sys
from pathlib import Path


APP_NAME = "PDF2MD"
APP_VERSION = "1.0.3"
APP_PUBLISHER = "PDF2MD"
DEFAULT_PORT = 5000
DATA_RETENTION_DAYS = 30


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def resource_root() -> Path:
    if is_frozen() and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def application_data_root() -> Path:
    override = os.environ.get("PDF2MD_DATA_DIR")
    if override:
        return Path(override).expanduser().resolve()
    if is_frozen():
        local_app_data = os.environ.get("LOCALAPPDATA")
        if local_app_data:
            return Path(local_app_data) / APP_NAME
    return Path(__file__).resolve().parent


DATA_DIR = application_data_root()
UPLOAD_DIR = DATA_DIR / "uploads"
OUTPUT_DIR = DATA_DIR / "output"
LOG_DIR = DATA_DIR / "logs"
RUNTIME_FILE = DATA_DIR / "runtime.json"


def ensure_runtime_directories() -> None:
    for directory in (DATA_DIR, UPLOAD_DIR, OUTPUT_DIR, LOG_DIR):
        directory.mkdir(parents=True, exist_ok=True)


def configure_tesseract() -> Path | None:
    """Locate bundled or locally installed Spanish/English OCR data."""
    configured = os.environ.get("TESSDATA_PREFIX")
    roots = [
        resource_root() / "tesseract",
        Path(sys.executable).resolve().parent / "tesseract",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Tesseract-OCR",
        Path(os.environ.get("PROGRAMFILES", "")) / "Tesseract-OCR",
        Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Tesseract-OCR",
    ]
    candidates = [Path(configured)] if configured else []
    candidates.extend(root / "tessdata" for root in roots)

    for candidate in candidates:
        required = (candidate / "spa.traineddata", candidate / "eng.traineddata")
        if candidate.is_dir() and all(item.is_file() for item in required):
            os.environ["TESSDATA_PREFIX"] = str(candidate)
            binary_dir = candidate.parent
            current_path = os.environ.get("PATH", "")
            if str(binary_dir).lower() not in current_path.lower():
                os.environ["PATH"] = str(binary_dir) + os.pathsep + current_path
            return candidate
    return None


ensure_runtime_directories()
