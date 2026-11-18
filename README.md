# Aptech Monthly BSR (Batch Status Report)

Desktop application that replaces the manual multi-sheet Excel “GLS Labstatus” workbook.

**One source of truth → every report updates automatically.**

## Features

- Central allocation form (enter data once)
- Automatic faculty workload calculation
- Automatic career / semester summaries + percentages
- Automatic MWF vs TTS batch & student totals
- Automatic STC / SP counts
- Faculty × Career matrix
- Visiting faculty list
- Lab capacity warnings
- Schedule conflict detection (same lab + time + day group)
- One-click Excel report export that follows the **GLS Labstatus template** structure
- Fully offline, single-file SQLite database

## Excel template fidelity

The official workbook `templates/GLS_Labstatus_template.xlsx` (sheet **August-26**) is the visual master.
`templates/14-August-2026 GLS Labstatus.xlsx` is a lightweight single-sheet copy of the same layout.

Export behaviour:

1. The entire Lab Status sheet is **copied** from the template (fonts, fills, merges, row heights, column widths, page setup).
2. **All sample / previously filled values are cleared** — the template is used strictly as a layout reference.
3. Live allocation data is written **value-only** into the exact same cells; no fonts or fills are overridden, so Arial Black / Arial 28 sizing and purple block fills stay intact.
4. Multiple allocations in the same lab/slot/day stack as successive career + student rows (matching the parent template).
5. Bottom tables (per-slot Career/STC/Total Batch, Faculty Load, Semester summary) are cleared then filled **in-place** at the template positions.

| Sheet | Purpose |
|-------|---------|
| **Lab Status** | Exact visual copy of the monthly GLS sheet, populated from live data |
| **Dashboard** | Career summary, MWF/TTS totals, faculty workload, faculty × career matrix |
| **Allocations** | Flat filterable list of every allocation |

## Quick Start

```bash
pip install -r requirements.txt
python seed_sample_data.py      # optional – loads demo rows from August-26 sheet
python main.py
```

## Refreshing the source workbook

For the Windows application, place the latest GLS Labstatus workbook in the `data` folder beside `AptechMonthlyBSR.exe`. The application discovers the newest `.xlsx` workbook in that folder and checks its content hash at startup. Replacing or updating the workbook triggers an import on the next launch; an unchanged workbook is not imported again. The same latest workbook is used as the export master template so its newest month sheet and historical tabs are carried into generated reports.

The default refresh policy is **Preserve application modifications**:

- New months and new allocations are added.
- Existing allocations are matched by report month, faculty, career, lab, time slot, day group and batch.
- For allocations imported by this version, source fields are refreshed only while their current values still match the last imported workbook values. Values changed inside the application are retained.
- Existing records created before import-baseline tracking are refreshed on the first merge only when their `created_at` and `updated_at` timestamps match; records previously edited in the application are preserved.
- Source rows are never used to delete allocations that are absent from a later workbook.
- A database backup is created before each changed-workbook import. Failed validation or import is rolled back.

The importer reads faculty names from each monthly sheet's roster and resolves full-name labels as well as initials. If the workbook layout is not recognized or no allocation rows can be parsed, the import is rejected instead of marking the workbook as successfully imported.

In packaged Windows builds, the database and generated reports use writable locations. The database uses the existing writable portable database when available; otherwise it is stored under `%LOCALAPPDATA%\\AptechMonthlyBSR\\database`. Generated reports use the app's `Output` folder when writable, otherwise `%LOCALAPPDATA%\\AptechMonthlyBSR\\Output`.

## How it works

1. **Lookups** (Labs, Faculty, Careers, Time Slots, Day Groups) are managed once.
2. **Allocations** are the only transactional data you enter. One teacher can have many allocations (multiple careers / batches) — each save creates a new row and never overwrites another.
3. Course/series (AI, PMTZ, DM …) is derived automatically from the batch code prefix (e.g. `AI-202412C2` → AI). There is no separate Course Title field.
4. Every dashboard number and Excel sheet is calculated live from the `allocations` table.

## Project layout

```
Aptech_LabTimetable/
├── main.py                 # Entry point
├── database.py             # SQLite layer (path → database/lab_status.db)
├── models.py               # CRUD including faculty management
├── calculations.py         # All aggregates & reports
├── export_excel.py         # Template-aligned Excel exporter
├── seed_sample_data.py     # Optional demo data
├── ui/
│   └── main_window.py      # Tkinter GUI (Allocations / Dashboard / Faculty)
├── templates/              # Master Excel templates (do not edit casually)
├── database/
│   └── lab_status.db       # Persistent SQLite file – backup this
├── Output/                 # Generated Lab Status Excel reports only
├── requirements.txt
└── README.md
```

### Database backup / restore
- The live database is a normal file: `database/monthly_bsr.db`
- **Backup:** copy that file somewhere safe
- **Restore:** stop the app, replace `database/monthly_bsr.db` with your backup, start again
- Startup never wipes an existing database; lookup seeds run only on empty tables

### Faculty management
Use the **Faculty** tab to Add / Edit / Deactivate / Reactivate / Delete teachers.
Inactive teachers stay in history but are hidden from new allocation dropdowns.
Prefer **Deactivate** over Delete when the teacher has past allocations.

## Building a Windows .exe

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
pyinstaller --noconfirm --windowed --name AptechTimetable main.py
```

## Notes on the template

- Monthly sheets in the official file (May-25 … August-26) use a dense merged-cell layout.
- The exporter recreates that layout programmatically so it stays correct when labs, slots or allocations change.
- The original template file is shipped under `templates/` for reference and future exact-clone enhancements.

---

Internal use – Aptech Computer Education / GLS.
