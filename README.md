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
- Import monthly BSR workbooks from the application (month-scoped replacement)
- Add and edit individual allocation records without editing Excel manually
- Automatically refresh the in-app Excel preview after allocation changes
- One-click Excel report export that follows the **GLS Labstatus template** structure
- Fully offline, single-file SQLite database

## Excel template fidelity

The official workbook `templates/21-August-2026 GLS Labstatus.xlsx` is the visual master included in this repository. The root-level `GLS_Labstatus_template.xlsx` is also retained as a reference copy.

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

## Import, edit and preview workflow

1. Choose **Import Excel** and select a monthly BSR workbook that follows the GLS Labstatus layout.
2. The importer recognizes month names in worksheet titles (for example, `August-26`, `August 2026`, or `21-August-2026`).
3. Confirm the import. Only allocation records for months present in the selected workbook are replaced; other months and lookup records are retained.
4. Select a month, then use **Add Allocation** or **Edit Selected** to maintain records in the application.
5. After an allocation is saved, the application regenerates the formatted workbook and refreshes the **Excel Preview** tab.
6. Use **Generate Monthly BSR** when you want to create the final report file in `%LOCALAPPDATA%\\AptechMonthlyBSR\\Output`.

The importer validates that it can read allocation rows before replacing a month's existing records. Keep a backup of `%LOCALAPPDATA%\\AptechMonthlyBSR\\database\\monthly_bsr.db` before importing important workbooks.

The current workbook contains two different faculty members with the initial `H`
(Hadiqa and Hasnain Arif). To avoid merging their allocations, the importer
keeps Hadiqa as `H` and maps Hasnain Arif to the separate `HA` roster entry.

## Standalone application and local data

The desktop application is the complete BSR system. External spreadsheet
synchronization is discontinued and is not required. The application does not
need a Google account or internet connection for normal operation.

- **SQLite is the only source of truth.** Allocation, faculty, career, module,
  lab, time-slot and reporting data is maintained by the application.
- Add, edit, deactivate and delete records inside the application.
- The dashboard and Excel preview are generated from the local database.
- Import an existing Excel BSR workbook to bring historical data into the app.
  After migration, maintain records in the application.
- Export a formatted monthly BSR workbook whenever an Excel copy is required.
- Application data is stored separately from the executable under
  `%LOCALAPPDATA%\\AptechMonthlyBSR`.
- Back up `%LOCALAPPDATA%\\AptechMonthlyBSR\\database\\monthly_bsr.db`
  regularly. Generated Excel reports are stored in
  `%LOCALAPPDATA%\\AptechMonthlyBSR\\Output`.

Excel is an import/export format only. It is not the live database and no
external spreadsheet is synchronized.

## Quick Start

```bash
pip install -r requirements.txt
python seed_sample_data.py      # optional – loads demo rows from August-26 sheet
python main.py
```

## How it works

1. **Lookups** (Labs, Faculty, Careers, Time Slots, Day Groups) are managed once.
2. **Allocations** are the only transactional data you enter. One teacher can have many allocations (multiple careers / batches) — each save creates a new row and never overwrites another.
3. Course/series (AI, PMTZ, DM …) is derived automatically from the batch code prefix (e.g. `AI-202412C2` → AI). There is no separate Course Title field.
4. Every dashboard number and Excel sheet is calculated live from the `allocations` table.

## Project layout

~~~~
AptechMonthlyBSR/
├── main.py                  # Application entry point
├── app_paths.py             # Bundled-resource and user-data paths
├── database.py              # SQLite database layer
├── models.py                # Data access and CRUD operations
├── calculations.py          # Dashboard and report calculations
├── export_excel.py          # Template-aligned Excel exporter
├── import_historical.py     # Excel import
├── ui/
│   ├── main_window.py       # Tkinter desktop interface
│   ├── excel_preview.py     # In-app workbook preview
│   └── splash.py
├── templates/               # Bundled master Excel templates
├── tests/                   # Regression tests
├── AptechMonthlyBSR.spec    # PyInstaller build definition
├── build_windows.bat        # Windows build script
├── requirements.txt
└── README.md
~~~~

Runtime data is stored separately under `%LOCALAPPDATA%\\AptechMonthlyBSR`:
SQLite database, imported master workbook, and generated Excel reports. These
files are not stored in or overwritten by the executable bundle.

### Database backup / restore
- Live database: `%LOCALAPPDATA%\\AptechMonthlyBSR\\database\\monthly_bsr.db`
- **Backup:** close the application and copy the database file to a safe location.
- **Restore:** close the application and replace the database file with your backup.
- Startup never wipes an existing database; lookup seeds run only on empty tables.
- Existing database and report files in the earlier project-folder layout are
  copied into the new data folder on first launch when possible.

### Faculty management
Use the **Faculty** tab to Add / Edit / Deactivate / Reactivate / Delete teachers.
Inactive teachers stay in history but are hidden from new allocation dropdowns.
Prefer **Deactivate** over Delete when the teacher has past allocations.

## Building the Windows application

Build on a Windows machine (the build is not cross-platform). Install Python
3.11 or newer, then run:

```bat
build_windows.bat
```

The script installs project dependencies and builds an **on-directory**
application with PyInstaller. The distributable application is created at:

```
dist\\AptechMonthlyBSR\\AptechMonthlyBSR.exe
```

Distribute the complete `dist\\AptechMonthlyBSR` folder (not the EXE by
itself). Zip that folder for convenient delivery. The master Excel template is
bundled with the application. User data and generated reports remain outside
the installation folder, so updating the application does not overwrite them.
On first launch, the app also copies an existing database, imported master
workbook, and generated reports from the earlier project-folder layout when
those files are present; it does not overwrite files already in the new location.

For a manual build:

```bat
python -m pip install -r requirements.txt
python -m PyInstaller --clean --noconfirm AptechMonthlyBSR.spec
```

## Notes on the template

- Monthly sheets in the official file (May-25 … August-26) use a dense merged-cell layout.
- The exporter recreates that layout programmatically so it stays correct when labs, slots or allocations change.
- The original template file is shipped under `templates/` for reference and future exact-clone enhancements.

---

Internal use – Aptech Computer Education / GLS.
