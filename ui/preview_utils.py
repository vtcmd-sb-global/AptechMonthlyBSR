"""Pure helpers for choosing the worksheet shown in the BSR preview."""

from datetime import datetime
from pathlib import Path
import re
from typing import List, Optional


def preferred_sheet_name(sheet_names: List[str], workbook_path: str | Path) -> Optional[str]:
    """Choose the report month from the output filename, then newest month tab."""
    names = list(sheet_names)
    name_set = set(names)

    # Generated reports are named BSR_YYYY-MM.xlsx. Prefer that exact month.
    match = re.search(r"BSR_(\d{4})-(\d{2})", Path(workbook_path).stem, re.IGNORECASE)
    if match:
        try:
            target = datetime(int(match.group(1)), int(match.group(2)), 1).strftime("%B-%y")
            if target in name_set:
                return target
        except ValueError:
            pass

    month_numbers = {
        "jan": 1, "january": 1, "feb": 2, "february": 2, "feburary": 2,
        "mar": 3, "march": 3, "apr": 4, "april": 4, "may": 5,
        "jun": 6, "june": 6, "jul": 7, "july": 7, "aug": 8, "august": 8,
        "sep": 9, "sept": 9, "september": 9, "oct": 10, "october": 10,
        "nov": 11, "november": 11, "dec": 12, "december": 12,
    }
    candidates = []
    for name in names:
        iso = re.search(r"(?<!\d)(20\d{2})[-_\s](0?[1-9]|1[0-2])(?!\d)", name.strip())
        if iso:
            candidates.append((int(iso.group(1)), int(iso.group(2)), name))
            continue
        match = re.search(r"([a-z]+)[\s_-]+(\d{2,4})(?:$|[^0-9])", name.strip(), re.IGNORECASE)
        if not match:
            continue
        month = month_numbers.get(match.group(1).lower())
        if not month:
            continue
        year = int(match.group(2))
        if year < 100:
            year += 2000
        candidates.append((year, month, name))
    if candidates:
        return max(candidates)[2]

    helpers = {"dashboard summary", "_data", "chart", "chart1", "allocations", "bsr"}
    return next(
        (name for name in names if name.strip().lower() not in helpers),
        names[0] if names else None,
    )
