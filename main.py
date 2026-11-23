#!/usr/bin/env python3
"""Aptech Monthly BSR – entry point."""

from pathlib import Path
import sys

ROOT = Path(__file__).parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from database import init_db
from models import list_allocations


def ensure_historical_data():
    """Import bundled historical data only when the database has no allocations."""
    try:
        # Never replace or re-import over records entered by an end user.
        if list_allocations(active_only=False):
            return
        from import_historical import import_historical, HIST_CANDIDATES
        src = None
        for c in HIST_CANDIDATES + [ROOT / "data" / "21-August-2026 GLS Labstatus.xlsx"]:
            if c.exists():
                src = c
                break
        if src:
            print(f"Importing historical BSR data from {src.name}…")
            summary = import_historical(src, replace_all=False)
            print(f"Imported {summary['total_allocations']} allocations across {len(summary['months'])} months.")
    except Exception as e:
        print("Historical import skipped:", e)


def main():
    init_db()
    ensure_historical_data()
    from ui.main_window import main as ui_main
    ui_main()


if __name__ == "__main__":
    main()
