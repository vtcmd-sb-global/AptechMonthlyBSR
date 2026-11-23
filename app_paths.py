"""Resolve bundled read-only resources and writable per-user application data."""
from __future__ import annotations
import os
import sys
from pathlib import Path

def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))

def resource_dir() -> Path:
    if is_frozen():
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    return Path(__file__).resolve().parent

RESOURCE_DIR = resource_dir()

def resource_path(*parts: str) -> Path:
    return RESOURCE_DIR.joinpath(*parts)

def _user_data_root() -> Path:
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA") or (Path.home() / "AppData" / "Local"))
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_DATA_HOME") or (Path.home() / ".local" / "share"))
    return base / "AptechMonthlyBSR"

APP_DATA_DIR = _user_data_root()
USER_DATA_DIR = APP_DATA_DIR / "user_data"
OUTPUT_DIR = APP_DATA_DIR / "Output"
