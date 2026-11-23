"""Resolve bundled read-only resources and writable per-user application data."""
from __future__ import annotations
import os
import shutil
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


def legacy_roots() -> tuple[Path, ...]:
    """Return locations used by older source-folder / PyInstaller releases."""
    if is_frozen():
        exe_dir = Path(sys.executable).resolve().parent
        return (exe_dir, exe_dir.parent.parent)
    return (RESOURCE_DIR,)

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


def migrate_legacy_user_files() -> None:
    """Copy old project-folder templates and reports into per-user data.

    Existing files in the new location always win; old files are never deleted.
    """
    for legacy_root in legacy_roots():
        old_template = legacy_root / "user_data" / "latest_imported_bsr.xlsx"
        new_template = USER_DATA_DIR / "latest_imported_bsr.xlsx"
        if old_template.is_file() and not new_template.exists():
            try:
                USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
                shutil.copy2(old_template, new_template)
            except OSError:
                pass

        old_output = legacy_root / "Output"
        if old_output.is_dir():
            try:
                for old_report in old_output.glob("BSR_*.xlsx"):
                    new_report = OUTPUT_DIR / old_report.name
                    if not new_report.exists():
                        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
                        shutil.copy2(old_report, new_report)
            except OSError:
                pass
