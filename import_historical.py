#!/usr/bin/env python3
"""Import historical monthly BSR sheets into the database (strict, clean)."""

from __future__ import annotations

import re
import shutil
from difflib import SequenceMatcher
from datetime import datetime, date
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import load_workbook

from database import init_db, db_session
from app_paths import USER_DATA_DIR, resource_path

HIST_CANDIDATES = [
    resource_path("templates", "21-August-2026 GLS Labstatus.xlsx"),
    resource_path("templates", "GLS_Labstatus_template.xlsx"),
    resource_path("GLS_Labstatus_template.xlsx"),
    resource_path("data", "21-August-2026 GLS Labstatus.xlsx"),
    Path("/home/workdir/attachments/21-August-2026 GLS Labstatus.xlsx"),
    Path("/home/workdir/attachments/GLS_Labstatus_template.xlsx"),
]

SLOT_DAY_COLS = {
    ("B", "MWF"): 2, ("B", "TTS"): 8,
    ("C", "MWF"): 14, ("C", "TTS"): 20,
    ("D", "MWF"): 26, ("D", "TTS"): 30,
    ("E", "MWF"): 36, ("E", "TTS"): 42,
    ("F", "MWF"): 48, ("F", "TTS"): 56,
    ("G", "MWF"): 62, ("G", "TTS"): 68,
}

MONTH_NAME_TO_NUM = {
    "jan": 1, "january": 1, "feb": 2, "february": 2, "feburary": 2,
    "mar": 3, "march": 3, "apr": 4, "april": 4, "may": 5,
    "jun": 6, "june": 6, "jul": 7, "july": 7, "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9, "oct": 10, "october": 10,
    "nov": 11, "november": 11, "dec": 12, "december": 12,
}

# Canonical faculty roster (static reference data from BSR)
CANONICAL_FACULTY = [
    ("MS", "Shahbaz", "V"),
    ("AA", "Aman Ali", "V"),
    ("HM", "Hafiz Moiz", "V"),
    ("T", "Talha Qureshi", "P"),
    ("A", "Aousaja", "P"),
    ("BF", "M Bilal Faroqui", "P"),
    ("RA", "Riaz Ahmed", "P"),
    ("MAR", "M Ashar Rehan", "P"),
    ("Z", "Syed Zaid", "V"),
    ("KA", "Kashif Ali", "P"),
    ("AW", "Abdul Wahab", "P"),
    ("H", "Hadiqa", "V"),
    ("AR", "AR", "V"),
    ("SAA", "SAA", "V"),
    ("AM", "AM", "V"),
    ("MA", "MA", "V"),
    ("MAK", "MAK", "V"),
    ("MT", "MT", "V"),
    ("AD", "AD", "V"),
    # The latest workbook lists both Hadiqa and Hasnain Arif as "H".
    # Keep Hadiqa on H and use this existing distinct roster code for Hasnain.
    ("HA", "Hasnain Arif", "V"),
    ("U", "U", "V"),
]

FACULTY_NAME_ALIASES = {
    "hasnain": "HA",
    "hasnain arif": "HA",
}

KNOWN_CAREERS = {
    "CPISM", "DISM", "HDSE I", "HDSE II", "ADSE I", "ADSE II",
    "CDMA", "AID", "CS", "STC", "SP", "SEM I", "OST",
    "HDSE1", "HDSE11", "ADSE1", "ADSE11", "HDSE 1", "HDSE 11",
    "ADSE 1", "ADSE 11",
}

CAREER_ALIASES = {
    "HDSE 1": "HDSE I", "HDSE1": "HDSE I", "HDSE 11": "HDSE II", "HDSE11": "HDSE II",
    "ADSE 1": "ADSE I", "ADSE1": "ADSE I", "ADSE 11": "ADSE II", "ADSE11": "ADSE II",
}


def sheet_to_year_month(name: str) -> Optional[str]:
    """Recognise common monthly worksheet labels without relying on one spelling."""
    n = (name or "").strip().lower()
    # ISO-style labels: 2026-08 / 2026_08 / 2026 08.
    m = re.search(r"(?<!\d)(20\d{2})[-_\s](0?[1-9]|1[0-2])(?!\d)", n)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}"
    # Month names or abbreviations: August-26, Aug 2026, February_26.
    m = re.search(r"([a-z]+)[\s_-]+(\d{2,4})(?:$|[^0-9])", n)
    if not m:
        return None
    mon = MONTH_NAME_TO_NUM.get(m.group(1))
    if not mon:
        return None
    yy = int(m.group(2))
    year = 2000 + yy if yy < 100 else yy
    return f"{year}-{mon:02d}"


def _cell_str(v) -> str:
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def _is_date_like(v) -> bool:
    if isinstance(v, (datetime, date)):
        return True
    s = _cell_str(v)
    if re.match(r"^\d{4}-\d{2}-\d{2}", s):
        return True
    if re.match(r"^\d{1,2}[-/][A-Za-z]{3}[-/]\d{2,4}$", s):
        return True
    if re.match(r"^\d{1,2}-[A-Z]{3,9}-\d{2}$", s, re.I):
        return True
    return False


def _fmt_date(v) -> Optional[str]:
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    # Some source workbooks store dates as raw Excel serial numbers.
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            from openpyxl.utils.datetime import from_excel
            converted = from_excel(v)
            if isinstance(converted, (datetime, date)):
                return converted.strftime("%Y-%m-%d")
        except Exception:
            pass
    s = str(v).strip()
    if not s or s.lower().startswith("admission"):
        return None
    for fmt in ("%Y-%m-%d", "%d-%b-%y", "%d-%b-%Y", "%d/%m/%Y", "%d-%B-%y"):
        try:
            return datetime.strptime(s[:20], fmt).strftime("%Y-%m-%d")
        except Exception:
            pass
    return None


def _to_int(v) -> int:
    try:
        if v is None or v == "" or v == "-":
            return 0
        return int(float(v))
    except Exception:
        return 0


def _parse_faculty_line(text) -> Tuple[str, str]:
    """Parse a BSR faculty header and return its displayed name and course.

    Recent workbooks sometimes use full names (for example, '(AI) Talha')
    instead of roster initials (for example, '(AI) T').
    """
    if text is None or _is_date_like(text):
        return "", ""
    t = _cell_str(text)
    if not t or t.lower().startswith("admission"):
        return "", ""
    if re.search(r"\d{4}", t) and "-" in t:
        return "", ""
    m = re.match(r"^\(([^)]+)\)\s*(.*?)\s*$", t)
    if not m:
        return "", ""
    course, person = m.group(1).strip(), m.group(2).strip()
    if not person or course.upper() in ("MWF", "TTS", "M/W/F", "T/T/S"):
        return "", ""
    return person.upper(), course


def _normalise_person_name(value: str) -> str:
    return " ".join(re.findall(r"[a-z]+", _cell_str(value).lower()))


def _resolve_faculty_initials(display_name: str, maps: dict) -> Optional[str]:
    """Resolve a header's initials or displayed name to the canonical roster.

    Exact initials/name matches are preferred. A unique close token match
    handles common shortened display names such as 'Moiz' or 'Amaan'.
    """
    raw = (display_name or "").strip().upper()
    if raw in maps["faculty"]:
        return raw

    normal = _normalise_person_name(raw)
    if not normal:
        return None
    names = maps.get("faculty_names", {})
    if normal in names:
        return names[normal]

    raw_tokens = normal.split()
    scored = []
    for full_name, initials in names.items():
        name_tokens = full_name.split()
        score = 0.0
        if set(raw_tokens) & set(name_tokens):
            score = 1.0
        else:
            score = max(
                (SequenceMatcher(None, a, b).ratio()
                 for a in raw_tokens for b in name_tokens),
                default=0.0,
            )
        if score >= 0.86:
            scored.append((score, initials))

    if not scored:
        return None
    best = max(score for score, _ in scored)
    matches = {initials for score, initials in scored if score == best}
    return next(iter(matches)) if len(matches) == 1 else None


def _norm_career(text: str) -> str:
    t = _cell_str(text).upper().replace("  ", " ")
    return CAREER_ALIASES.get(t, t)


def _is_career(text) -> bool:
    """Recognise known programs and reasonable custom career labels.

    The previous allow-list silently discarded every institute-specific or
    newly introduced program that was not already hard-coded in the app.
    """
    if text is None or _is_date_like(text):
        return False
    raw = _cell_str(text)
    t = _norm_career(raw)
    if t in KNOWN_CAREERS or t in CAREER_ALIASES:
        return True
    if not raw or _is_batch_code(raw) or re.fullmatch(r"[-+]?\d+(?:\.\d+)?", raw):
        return False
    reserved = {
        "ADMISSION OPEN", "INITIAL", "STATUS", "FACULTY", "TOTAL",
        "TOTAL STUDENT", "TOTAL STUDENTS", "STC BEG", "STC END",
        "SP BEG", "SP END", "MWF", "TTS", "CI5", "C15", "CIS",
        "LAB", "MODULE", "BATCH", "STUDENTS",
    }
    if t in reserved or re.fullmatch(r"L[1-8]", t):
        return False
    return bool(re.fullmatch(r"[A-Z][A-Z0-9 &/+()._-]{0,29}", t))


def _is_batch_code(text) -> bool:
    if text is None or _is_date_like(text):
        return False
    s = _cell_str(text)
    # e.g. AI-202412C2, PMTZ-202405B, DM-202601C
    return bool(re.match(r"^[A-Za-z]{1,6}-\d{4,}", s))


def extract_faculty_from_sheet(ws) -> List[Tuple[str, str, str]]:
    """Read the faculty roster wherever it appears in the worksheet."""
    found = []
    header_row = None
    for r in range(1, (ws.max_row or 1) + 1):
        v = ws.cell(r, 2).value
        if v and str(v).strip().upper() == "INITIAL":
            header_row = r
            break
    if not header_row:
        return found
    for r in range(header_row + 1, min((ws.max_row or header_row) + 1, header_row + 100)):
        init = ws.cell(r, 2).value
        status = ws.cell(r, 4).value
        name = ws.cell(r, 5).value
        if not init:
            continue
        init_s = str(init).strip().upper()
        if not re.fullmatch(r"[A-Z]{1,5}", init_s):
            continue
        if _is_date_like(init) or _is_date_like(name):
            continue
        name_s = _cell_str(name) or init_s
        st = str(status or "V").strip().upper()[:1]
        if st not in ("P", "V"):
            st = "V"
        found.append((init_s, name_s, st))
    return found


def load_grid(ws, max_r=None, max_c=None) -> Dict[Tuple[int, int], Any]:
    """Load all populated cells; do not truncate imported workbooks at 100x75."""
    grid = {}
    max_r = max_r or (ws.max_row or 1)
    max_c = max_c or (ws.max_column or 1)
    for r_idx, row in enumerate(
        ws.iter_rows(min_row=1, max_row=max_r, max_col=max_c, values_only=True), 1
    ):
        for c_idx, value in enumerate(row, 1):
            if value not in (None, ""):
                grid[(r_idx, c_idx)] = value
    return grid



def _faculty_header_text(grid: dict, row: int, start_col: int) -> str:
    """Join split header cells without reading into the adjacent schedule block."""
    # Most BSR blocks are six columns wide, but the D-slot MWF block is only
    # four columns wide: its TTS neighbour starts at column 30. Derive each
    # block's true right edge from the next configured block start.
    next_starts = [col for col in SLOT_DAY_COLS.values() if col > start_col]
    end_col = min(next_starts) - 1 if next_starts else start_col + 5
    parts = []
    for col in range(start_col, end_col + 1):
        value = grid.get((row, col))
        text = _cell_str(value)
        if text:
            parts.append(text)
    return " ".join(parts)


def parse_sheet_allocations(ws, year_month: str, maps: dict, conn) -> List[dict]:
    """Parse by finding faculty header cells, then reading the block below them."""
    grid = load_grid(ws)
    # lab row markers
    lab_at = {}
    for (r, c), v in grid.items():
        if c == 1 and isinstance(v, str) and re.fullmatch(r"L[1-8]", v.strip().upper()):
            lab_at[r] = v.strip().upper()

    def lab_for_row(r: int) -> str:
        best = None
        for lr, code in lab_at.items():
            if lr <= r + 2:  # L marker can be on module row
                if best is None or lr > best[0]:
                    best = (lr, code)
        return best[1] if best else "L1"

    results = []
    # Scan every worksheet row; real BSRs can extend below the old row-54 cutoff.
    last_row = max((r for r, _c in grid.keys()), default=1)
    for (slot, day), start_c in SLOT_DAY_COLS.items():
        for r in range(1, last_row + 1):
            fac_txt = grid.get((r, start_c))
            admission_header = _cell_str(fac_txt).lower().startswith("admission")
            displayed_name, _hint = _parse_faculty_line(
                _faculty_header_text(grid, r, start_c)
            )
            if not admission_header and not displayed_name:
                continue
            initials = (
                _resolve_faculty_initials(displayed_name, maps)
                if displayed_name else None
            )
            if not admission_header and not initials:
                continue

            # Block rows relative to faculty row
            batch_r, mod_r, date_r = r + 1, r + 2, r + 3
            career_start = r + 4
            batch_raw = grid.get((batch_r, start_c))
            batch = None
            if batch_raw and not _is_date_like(batch_raw):
                bs = _cell_str(batch_raw)
                if re.search(r"[A-Za-z]", bs) and not bs.lower().startswith("admission"):
                    batch = bs

            module = _cell_str(grid.get((mod_r, start_c)))
            if module.lower().startswith("admission") or _is_date_like(grid.get((mod_r, start_c))):
                module = ""
            # If module row is actually an L marker in col A only, module is still this cell
            mod_clean = re.sub(r"\s+\d+$", "", module).strip() if module else ""
            if re.fullmatch(r"L[1-7]", mod_clean.upper()):
                mod_clean = ""

            mod_start = _fmt_date(grid.get((date_r, start_c)))
            admission = admission_header
            if _cell_str(grid.get((date_r, start_c))).lower().startswith("admission"):
                admission = True

            careers_found = []
            blank_career_rows = 0
            # Read the complete career list in this block. The old four-row
            # limit dropped additional programs from otherwise valid batches.
            for cr in range(career_start, min(last_row + 1, career_start + 16)):
                if _parse_faculty_line(_faculty_header_text(grid, cr, start_c))[0]:
                    break
                cval = grid.get((cr, start_c))
                if cval in (None, ""):
                    blank_career_rows += 1
                    if blank_career_rows >= 2:
                        break
                    continue
                blank_career_rows = 0
                if not _is_career(cval):
                    continue
                stud = _to_int(grid.get((cr, start_c + 4)))
                if stud == 0:
                    stud = _to_int(grid.get((cr, start_c + 3)))
                careers_found.append((_norm_career(_cell_str(cval)), stud))

            # batch start: next date after careers or CI5 row nearby
            batch_start = None
            for br in range(career_start, career_start + 6):
                if str(grid.get((br, 1)) or "").strip().upper() in ("CI5", "C15", "CIS"):
                    batch_start = _fmt_date(grid.get((br, start_c)))
                    break
            if not batch_start:
                for br in range(date_r + 1, date_r + 6):
                    if _is_date_like(grid.get((br, start_c))):
                        batch_start = _fmt_date(grid.get((br, start_c)))
                        break

            if not careers_found and not admission:
                continue

            lab_code = lab_for_row(r)
            fid = ensure_faculty(conn, initials, maps, displayed_name) if displayed_name else None
            if fid:
                initials = next(
                    (code for code, faculty_id in maps["faculty"].items() if faculty_id == fid),
                    initials,
                )
            slot_id = maps["slots"].get(slot)
            day_id = maps["days"].get(day)
            lab_id = maps["labs"].get(lab_code)
            if not all([slot_id, day_id, lab_id]):
                continue

            # An empty Admission Open block is still a real BSR state. Keep
            # it editable in the application instead of silently dropping it.
            if not careers_found and admission:
                results.append({
                    "faculty_id": fid,
                    "career_id": None,
                    "lab_id": lab_id,
                    "time_slot_id": slot_id,
                    "day_group_id": day_id,
                    "batch_code": batch,
                    "module_name": mod_clean or None,
                    "students": 0,
                    "module_start_date": mod_start,
                    "actual_start_date": batch_start,
                    "notes": None,
                    "is_admission_open": True,
                    "report_month": year_month,
                    "career_code": "",
                    "faculty_initials": initials or "",
                })
                continue

            for career_code, students in careers_found:
                cid = ensure_career(conn, career_code, maps)
                if not cid:
                    continue
                results.append({
                    "faculty_id": fid,
                    "career_id": cid,
                    "lab_id": lab_id,
                    "time_slot_id": slot_id,
                    "day_group_id": day_id,
                    "batch_code": batch,
                    "module_name": mod_clean or None,
                    "students": students,
                    "module_start_date": mod_start,
                    "actual_start_date": batch_start,
                    "notes": None,
                    "is_admission_open": admission,
                    "report_month": year_month,
                    "career_code": career_code,
                    "faculty_initials": initials,
                })
    return results


def parse_sheet_metrics(ws, year_month: str) -> dict:
    metrics = {
        "year_month": year_month,
        "stc_beg": 0, "stc_end": 0, "sp_beg": 0, "sp_end": 0,
        "total_batches": 0, "total_students": 0,
        "mwf_batches": 0, "mwf_students": 0,
        "tts_batches": 0, "tts_students": 0,
        "report_path": None, "notes": "imported from historical Excel",
    }
    rows = list(ws.iter_rows(
        min_row=min(55, ws.max_row or 55),
        max_row=ws.max_row or 100,
        max_col=ws.max_column or 75,
        values_only=True,
    ))
    for row in rows:
        for c, v in enumerate(row):
            if v is None:
                continue
            s = str(v).strip().lower()

            def right_num():
                for oc in range(c + 1, min(c + 8, len(row))):
                    nv = row[oc]
                    if nv is not None and nv != "":
                        try:
                            return int(float(nv))
                        except Exception:
                            pass
                return 0

            if "stc beg" in s:
                metrics["stc_beg"] = right_num()
            elif "stc end" in s:
                metrics["stc_end"] = right_num()
            elif "sp beg" in s:
                metrics["sp_beg"] = right_num()
            elif "sp end" in s:
                metrics["sp_end"] = right_num()
            elif "total batch mwf" in s:
                metrics["mwf_batches"] = right_num()
            elif "total batch tts" in s:
                metrics["tts_batches"] = right_num()
            elif s in ("total student", "total students"):
                n = right_num()
                if c >= 55 and c <= 62:
                    metrics["mwf_students"] = n
                elif c >= 63:
                    metrics["tts_students"] = n
            elif "stc total" in s:
                metrics["stc_end"] = metrics["stc_end"] or right_num()
    metrics["total_batches"] = metrics["mwf_batches"] + metrics["tts_batches"]
    if not metrics["total_students"]:
        metrics["total_students"] = metrics["mwf_students"] + metrics["tts_students"]
    return metrics


def ensure_faculty(conn, initials: str, maps: dict, display_name: str = "") -> Optional[int]:
    """Resolve a roster member or create a missing faculty member from the sheet.

    Unknown faculty names must not cause their allocation blocks to disappear.
    Generated initials are deterministic and kept unique within the roster.
    """
    key = (initials or "").upper()
    if key and key in maps["faculty"]:
        return maps["faculty"][key]

    name = _cell_str(display_name).strip()
    normal = _normalise_person_name(name)
    if normal and normal in maps.get("faculty_names", {}):
        resolved = maps["faculty_names"][normal]
        return maps["faculty"].get(resolved)

    if not name:
        return None
    tokens = normal.split()
    letters = "".join(tokens).upper() or re.sub(r"[^A-Z]", "", name.upper())
    base = "".join(token[0] for token in tokens).upper() or "F"
    candidates = [base[:5]]
    candidates.extend((base + letters[i])[:5] for i in range(len(letters)))
    candidates.extend(("F" + letters[i:i + 4])[:5] for i in range(len(letters)))
    unique = next((candidate for candidate in candidates if candidate and candidate not in maps["faculty"]), None)
    if not unique:
        for suffix in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            candidate = (base[:4] + suffix)[:5]
            if candidate not in maps["faculty"]:
                unique = candidate
                break
    if not unique:
        return None

    cur = conn.execute(
        "INSERT INTO faculty (initials, full_name, status, active) VALUES (?, ?, 'V', 1)",
        (unique, name),
    )
    maps["faculty"][unique] = cur.lastrowid
    maps.setdefault("faculty_names", {})[normal] = unique
    return cur.lastrowid


def ensure_career(conn, code: str, maps: dict) -> Optional[int]:
    key = _norm_career(code)
    if not key:
        return None
    if key in maps["careers"]:
        return maps["careers"][key]
    is_stc = 1 if key == "STC" else 0
    is_sp = 1 if key == "SP" else 0
    cur = conn.execute(
        "INSERT INTO careers (code, name, is_stc, is_sp, sort_order) VALUES (?, ?, ?, ?, 150)",
        (key, key, is_stc, is_sp),
    )
    maps["careers"][key] = cur.lastrowid
    return cur.lastrowid


def seed_canonical_faculty(conn) -> None:
    existing = {r[0].upper() for r in conn.execute("SELECT initials FROM faculty").fetchall()}
    for init, name, status in CANONICAL_FACULTY:
        if init.upper() in existing:
            # fix name if currently equals initials or looks like a date
            conn.execute(
                """
                UPDATE faculty SET full_name = ?, status = ?, active = 1
                WHERE UPPER(initials) = ?
                  AND (full_name = initials OR full_name GLOB '*[0-9]*' OR full_name = 'XX')
                """,
                (name, status, init.upper()),
            )
            continue
        conn.execute(
            "INSERT INTO faculty (initials, full_name, status, active) VALUES (?, ?, ?, 1)",
            (init.upper(), name, status),
        )


def purge_bad_data(conn) -> None:
    """Remove date-like faculty, XX faculty, and allocations pointing at them."""
    bad_ids = [
        r[0] for r in conn.execute(
            """
            SELECT id FROM faculty WHERE
                initials = 'XX' OR full_name = 'XX'
                OR initials GLOB '*[0-9]*'
                OR full_name GLOB '[0-9]*'
                OR length(initials) > 6
                OR UPPER(initials) IN ('HTML','PHP','JAVA','SQL','SEO','GIT','FSA','MUI','PWD','STC','SP','WORDPRESS')
            """
        ).fetchall()
    ]
    if bad_ids:
        placeholders = ",".join("?" * len(bad_ids))
        conn.execute(f"DELETE FROM allocations WHERE faculty_id IN ({placeholders})", bad_ids)
        conn.execute(f"DELETE FROM faculty WHERE id IN ({placeholders})", bad_ids)

    # also delete batch-code-looking faculty
    bad2 = [
        r[0] for r in conn.execute(
            "SELECT id FROM faculty WHERE initials LIKE '%-%' OR full_name LIKE '%-%'"
        ).fetchall()
    ]
    if bad2:
        ph = ",".join("?" * len(bad2))
        conn.execute(f"DELETE FROM allocations WHERE faculty_id IN ({ph})", bad2)
        conn.execute(f"DELETE FROM faculty WHERE id IN ({ph})", bad2)


def import_historical(path: Optional[Path] = None, replace_all: bool = True) -> dict:
    init_db()
    src = path
    if src is None:
        for c in HIST_CANDIDATES:
            if c.exists():
                src = c
                break
    if src is None or not Path(src).exists():
        raise FileNotFoundError("Historical BSR Excel not found")
    src = Path(src)
    wb = load_workbook(src, data_only=True)

    summary = {"file": str(src), "months": {}, "total_allocations": 0}

    with db_session() as conn:
        if replace_all:
            conn.execute("DELETE FROM allocations")
            purge_bad_data(conn)
        seed_canonical_faculty(conn)

        # merge faculty names from latest sheets
        for sheet_name in reversed(wb.sheetnames):
            if not sheet_to_year_month(sheet_name):
                continue
            for init, name, status in extract_faculty_from_sheet(wb[sheet_name]):
                row = conn.execute(
                    "SELECT id, full_name FROM faculty WHERE UPPER(initials)=?",
                    (init,),
                ).fetchone()
                if row:
                    if row[1] == init or not row[1]:
                        conn.execute(
                            "UPDATE faculty SET full_name=?, status=?, active=1 WHERE id=?",
                            (name, status, row[0]),
                        )
                else:
                    conn.execute(
                        "INSERT INTO faculty (initials, full_name, status, active) VALUES (?,?,?,1)",
                        (init, name, status),
                    )

        faculty_rows = conn.execute(
            "SELECT id, initials, full_name FROM faculty"
        ).fetchall()
        maps = {
            "faculty": {r["initials"].upper(): r["id"] for r in faculty_rows},
            "faculty_names": {
                **{
                    _normalise_person_name(r["full_name"]): r["initials"].upper()
                    for r in faculty_rows if _normalise_person_name(r["full_name"])
                },
                **FACULTY_NAME_ALIASES,
            },
            "labs": {r[0].upper(): r[1] for r in conn.execute("SELECT lab_code, id FROM labs").fetchall()},
            "careers": {r[0].upper(): r[1] for r in conn.execute("SELECT code, id FROM careers").fetchall()},
            "slots": {r[0].upper(): r[1] for r in conn.execute("SELECT code, id FROM time_slots").fetchall()},
            "days": {r[0].upper(): r[1] for r in conn.execute("SELECT code, id FROM day_groups").fetchall()},
        }
        for i in range(1, 9):
            code = f"L{i}"
            if code not in maps["labs"]:
                cur = conn.execute("INSERT INTO labs (lab_code, pcs, active) VALUES (?,20,1)", (code,))
                maps["labs"][code] = cur.lastrowid

        for sheet_name in wb.sheetnames:
            ym = sheet_to_year_month(sheet_name)
            if not ym:
                continue
            ws = wb[sheet_name]
            # Parse and validate before replacing anything for this month. A
            # renamed/unsupported or empty worksheet must never erase good data.
            allocs = parse_sheet_allocations(ws, ym, maps, conn)
            if not allocs:
                raise ValueError(
                    f"No allocation records could be read from worksheet "
                    f"{sheet_name!r}. Existing data has been preserved."
                )
            # Import is month-scoped: other months remain untouched.
            conn.execute("DELETE FROM allocations WHERE report_month=?", (ym,))
            for a in allocs:
                conn.execute(
                    """
                    INSERT INTO allocations (
                        faculty_id, career_id, lab_id, time_slot_id, day_group_id,
                        batch_code, course_title, module_name, students,
                        module_start_date, actual_start_date, notes, is_admission_open,
                        report_month, is_active
                    ) VALUES (?,?,?,?,?,?,NULL,?,?,?,?,?,?,?,1)
                    """,
                    (
                        a["faculty_id"], a["career_id"], a["lab_id"],
                        a["time_slot_id"], a["day_group_id"],
                        a.get("batch_code"), a.get("module_name"), a.get("students", 0),
                        a.get("module_start_date"), a.get("actual_start_date"),
                        a.get("notes"), int(a.get("is_admission_open") or 0), ym,
                    ),
                )
            metrics = parse_sheet_metrics(ws, ym)
            if not metrics.get("total_students"):
                metrics["total_students"] = sum(x.get("students", 0) for x in allocs)
            if not metrics.get("total_batches"):
                metrics["total_batches"] = len(allocs)
            stc_stu = sum(x["students"] for x in allocs if x.get("career_code") == "STC")
            sp_stu = sum(x["students"] for x in allocs if x.get("career_code") == "SP")
            metrics["stc_end"] = metrics.get("stc_end") or stc_stu
            metrics["sp_end"] = metrics.get("sp_end") or sp_stu
            metrics["stc_beg"] = metrics.get("stc_beg") or metrics["stc_end"]
            metrics["sp_beg"] = metrics.get("sp_beg") or metrics["sp_end"]

            existing = conn.execute(
                "SELECT id FROM monthly_bsr_metrics WHERE year_month=?", (ym,)
            ).fetchone()
            fields = (
                "stc_beg", "stc_end", "sp_beg", "sp_end",
                "total_batches", "total_students",
                "mwf_batches", "mwf_students", "tts_batches", "tts_students",
                "report_path", "notes",
            )
            vals = [metrics.get(f, 0) for f in fields]
            if existing:
                sets = ", ".join(f"{f}=?" for f in fields)
                conn.execute(
                    f"UPDATE monthly_bsr_metrics SET {sets} WHERE year_month=?",
                    (*vals, ym),
                )
            else:
                cols = ", ".join(["year_month"] + list(fields))
                ph = ", ".join(["?"] * (1 + len(fields)))
                conn.execute(
                    f"INSERT INTO monthly_bsr_metrics ({cols}) VALUES ({ph})",
                    (ym, *vals),
                )

            summary["months"][ym] = {
                "sheet": sheet_name,
                "allocations": len(allocs),
                "students": metrics.get("total_students", 0),
            }
            summary["total_allocations"] += len(allocs)

        # final purge of any residual garbage faculty
        purge_bad_data(conn)
        # remove faculty names that are career codes / modules if no valid name
        conn.execute(
            """
            DELETE FROM faculty WHERE id NOT IN (SELECT DISTINCT faculty_id FROM allocations)
              AND UPPER(initials) IN (
                'HTML','PHP','JAVA','SQL','SEO','GIT','FSA','MUI','PWD','STC','SP',
                'WORDPRESS','CDMA','CPISM','DISM','OST','AI','MWF','TTS','FBDS',
                'MYSQL','DART','VAT','ENT','PBD'
              )
            """
        )

    wb.close()

    # Keep the uploaded workbook as the next export's visual master. This
    # preserves the newest monthly sheet layout, merges, widths and styling.
    imported_template = USER_DATA_DIR / "latest_imported_bsr.xlsx"
    imported_template.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, imported_template)

    return summary


if __name__ == "__main__":
    s = import_historical()
    print("Imported", s["total_allocations"], "allocations")
    for ym, info in sorted(s["months"].items()):
        print(f"  {ym}: {info}")
