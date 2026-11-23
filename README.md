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
6. Use **Generate Monthly BSR** when you want to create the final report file in `Output/`.

The importer validates that it can read allocation rows before replacing a month's existing records. Keep a backup of `database/monthly_bsr.db` before importing important workbooks.

The current workbook contains two different faculty members with the initial `H`
(Hadiqa and Hasnain Arif). To avoid merging their allocations, the importer
keeps Hadiqa as `H` and maps Hasnain Arif to the separate `HA` roster entry.

## Google Sheets synchronization (optional)

The desktop application can update an existing Google Sheets BSR tab without
rebuilding its formatting. It sends the generated worksheet's cell values to a
small Google Apps Script bridge. The bridge clears **values only** and writes
the updated values into the existing tab; cell formatting, merged ranges,
column widths, row heights and validations remain in Google Sheets.

### One-time setup

1. Open the institute's Google spreadsheet using an account authorized to edit it.
2. Open **Extensions → Apps Script** and paste the contents of
   `google_apps_script/Code.gs` into the script editor.
3. In **Project Settings → Script Properties**, add both properties:
   - `BSR_SYNC_TOKEN`: a long, randomly generated secret;
   - `BSR_SPREADSHEET_ID`: `1T8sPu5d7ntMNqv79N3rCiM2fJCyC1qYK94eU0RY8Bs8`.
   Keep the token private and do not commit it to GitHub.
4. Deploy the script as a **Web app**, executing as the spreadsheet owner.
   Prefer restricting access to your institute's Google Workspace domain where
   available. If the deployment requires public access, the secret is still
   required for every write request; protect the deployment URL and secret.
5. In the desktop application, choose **Google Sheets Sync** and enter:
   - the deployed Web App URL;
   - the spreadsheet ID from its URL;
   - the target worksheet's numeric `gid` as a fallback (the supplied link
     currently points to `1463201093`);
   - the same sync secret configured in Script Properties.
6. Select the report month and click **Save & Sync** once. This saves the
   connection settings and performs the first sync. The app generates the
   selected month's workbook, then sends that worksheet to the bridge. The
   bridge first looks for a Google Sheets tab with the matching month name
   (for example, `November-26`); if no matching tab exists, it uses the
   configured `gid` fallback.

**Automatic synchronization after setup**

After the first successful **Save & Sync**, synchronization runs automatically
for subsequent application changes. The desktop application queues the affected
report month(s), regenerates each formatted BSR workbook, and publishes the
worksheet values through the Apps Script bridge in the background.

Automatic sync is triggered after:
- Adding, editing, deactivating or deleting an allocation;
- Clearing allocations;
- Importing a workbook (each imported month is queued);
- Generating a monthly BSR report;
- Adding, editing, deactivating, reactivating or deleting faculty;
- Adding, editing or deleting a career/course.

The application keeps local SQLite as the source of truth. Google Sheets is a
reporting copy; edits made directly in Google Sheets are not imported back into
the application.

**Failure recovery and status**

- A failed or offline sync does not undo or block a successful local database
  change.
- Pending report months are saved in
  `user_data/google_sheets_sync_state.json` and survive application restarts.
- The application retries failed synchronization while it remains open and
  retries pending months again on the next launch.
- The bottom status bar reports synchronization progress, pending months and
  the time of the last successful sync.
- Use **Google Sheets Sync** to configure the bridge, perform a manual sync of
  the selected month, or retry a pending report.

**Important:** sync replaces cell values on the selected month tab only. It
does not change formatting, merged cells, column widths, row heights,
validations or other tabs. The desktop app does not access or modify the
institute's spreadsheet until an authorized user configures the bridge.
Keep the sync secret private. The desktop app stores its configuration locally
in `user_data/google_sheets.json`; protect access to the Windows account and
never commit either user-data file to GitHub.

The bridge is optional. Excel import, editing, preview and downloaded workbook
generation continue to work without Google credentials.

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
