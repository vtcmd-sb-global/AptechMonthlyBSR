#!/usr/bin/env python3
"""Aptech Monthly BSR – entry point."""

from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from database import init_db
from paths import find_source_workbook


def ensure_historical_data():
    """Merge the newest source workbook on startup when its contents change."""
    try:
        from import_historical import import_historical
        src = find_source_workbook()
        if not src:
            print("No source BSR workbook found in the data folder.")
            return
        summary = import_historical(src, replace_all=False)
        if summary.get("skipped"):
            print(f"Source workbook unchanged: {src.name}")
        else:
            print(
                f"Imported {summary.get('inserted_allocations', 0)} new allocations "
                f"across {len(summary.get('months', {}))} months; "
                f"preserved {summary.get('skipped_duplicates', 0)} existing allocations."
            )
    except Exception as e:
        print("Historical import skipped:", e)


def main():
    init_db()
    ensure_historical_data()
    from ui.main_window import main as ui_main
    ui_main()


if __name__ == "__main__":
    main()
