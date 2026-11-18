"""Resolve bundled resources and writable user data for source and frozen builds."""
from __future__ import annotations

import os
import sys
from pathlib import Path

FROZEN = bool(getattr(sys, "frozen", False))
APP_ROOT = Path(sys.executable).resolve().parent if FROZEN else Path(__file__).resolve().parent
RESOURCE_ROOT = Path(getattr(sys, "_MEIPASS", APP_ROOT)).resolve()

_LOCAL_BASE = Path(os.environ.get("LOCALAPPDATA") or (Path.home() / "AppData" / "Local"))
USER_DATA_ROOT = _LOCAL_BASE / "AptechMonthlyBSR"


def _can_write(directory: Path) -> bool:
    try:
        directory.mkdir(parents=True, exist_ok=True)
        probe = directory / ".write_test"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink(missing_ok=True)
        return True
    except OSError:
        return False


def get_database_dir() -> Path:
    # Use the portable folder whenever writable; Program Files falls back to AppData.
    for candidate in (APP_ROOT / "database", RESOURCE_ROOT / "database"):
        if _can_write(candidate):
            return candidate
    return USER_DATA_ROOT / "database"


def get_output_dir() -> Path:
    portable = APP_ROOT / "Output"
    return portable if _can_write(portable) else USER_DATA_ROOT / "Output"


def get_data_dirs() -> list[Path]:
    result = []
    for candidate in (APP_ROOT / "data", RESOURCE_ROOT / "data"):
        if candidate not in result and candidate.is_dir():
            result.append(candidate)
    return result


def find_source_workbook() -> Path | None:
    """Choose the newest Excel workbook in the external data folder(s)."""
    candidates = []
    for directory in get_data_dirs():
        for path in directory.glob("*.xlsx"):
            if path.name.startswith("~$") or path.name.startswith("."):
                continue
            candidates.append(path)
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


def get_template_candidates() -> list[Path]:
    names = ("GLS_Labstatus_template.xlsx", "21-August-2026 GLS Labstatus.xlsx")
    result = []
    # The current source workbook is also the best master template because it
    # carries the latest month sheet and the full historical workbook structure.
    source = find_source_workbook()
    if source:
        result.append(source)
    for directory in (APP_ROOT / "templates", RESOURCE_ROOT / "templates", APP_ROOT / "data"):
        for name in names:
            path = directory / name
            if path not in result:
                result.append(path)
    # Also accept a single workbook template with a different filename.
    for directory in (APP_ROOT / "templates", RESOURCE_ROOT / "templates"):
        if directory.is_dir():
            result.extend(p for p in sorted(directory.glob("*.xlsx")) if p not in result)
    return result
