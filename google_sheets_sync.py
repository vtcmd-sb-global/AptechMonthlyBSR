"""Optional Google Sheets sync through a user-deployed Apps Script web app.

The desktop app never connects to or edits the spreadsheet until an operator
configures the bridge and explicitly selects Google Sheets Sync.
"""
from __future__ import annotations

import json
import urllib.request
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

ROOT = Path(__file__).resolve().parent
CONFIG_PATH = ROOT / "user_data" / "google_sheets.json"

SYNC_STATE_PATH = ROOT / "user_data" / "google_sheets_sync_state.json"


def _load_sync_state() -> dict:
    """Load the durable automatic-sync queue and last-sync metadata."""
    if not SYNC_STATE_PATH.exists():
        return {"pending_months": [], "last_success_at": None, "last_success_month": None, "last_error": None}
    try:
        state = json.loads(SYNC_STATE_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        state = {}
    pending = state.get("pending_months", [])
    if not isinstance(pending, list):
        pending = []
    return {
        "pending_months": sorted({str(month) for month in pending if _valid_month(str(month))}),
        "last_success_at": state.get("last_success_at"),
        "last_success_month": state.get("last_success_month"),
        "last_error": state.get("last_error"),
    }


def _valid_month(value: str) -> bool:
    try:
        datetime.strptime(value, "%Y-%m")
        return len(value) == 7
    except (TypeError, ValueError):
        return False


def _save_sync_state(state: dict) -> None:
    SYNC_STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    temp_path = SYNC_STATE_PATH.with_suffix(".tmp")
    temp_path.write_text(json.dumps(state, indent=2), encoding="utf-8")
    temp_path.replace(SYNC_STATE_PATH)


def get_google_sheets_sync_state() -> dict:
    """Return a safe copy of pending months and synchronization metadata."""
    return _load_sync_state()


def queue_google_sheets_sync(months) -> list[str]:
    """Persist months that need publishing; no-op until Sheets is configured."""
    config = load_google_sheets_config()
    required = ("web_app_url", "spreadsheet_id", "sheet_id", "token")
    if any(not str(config.get(key, "")).strip() for key in required):
        return []
    state = _load_sync_state()
    valid = {str(month) for month in months if _valid_month(str(month))}
    state["pending_months"] = sorted(set(state["pending_months"]) | valid)
    state["last_error"] = None
    _save_sync_state(state)
    return state["pending_months"]


def mark_google_sheets_month_synced(month: str) -> None:
    state = _load_sync_state()
    state["pending_months"] = [item for item in state["pending_months"] if item != month]
    state["last_success_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
    state["last_success_month"] = month
    state["last_error"] = None
    _save_sync_state(state)


def record_google_sheets_sync_error(error: str) -> None:
    state = _load_sync_state()
    state["last_error"] = str(error)
    _save_sync_state(state)



def load_google_sheets_config() -> dict:
    if not CONFIG_PATH.exists():
        return {}
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def save_google_sheets_config(config: dict) -> None:
    required = ("web_app_url", "spreadsheet_id", "sheet_id", "token")
    missing = [key for key in required if not str(config.get(key, "")).strip()]
    if missing:
        raise ValueError("Missing Google Sheets settings: " + ", ".join(missing))
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(
        json.dumps({key: str(config[key]).strip() for key in required}, indent=2),
        encoding="utf-8",
    )


def _serialise_cell(value: Any) -> Any:
    if isinstance(value, datetime):
        return {"type": "date", "value": value.strftime("%Y-%m-%dT%H:%M:%S")}
    if isinstance(value, date):
        return {"type": "date", "value": value.strftime("%Y-%m-%dT00:00:00")}
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def sync_workbook_to_google_sheet(
    workbook_path: str | Path,
    sheet_name: str,
    config: dict | None = None,
) -> dict:
    """Replace values on one existing Google Sheet tab, retaining its formatting."""
    config = config or load_google_sheets_config()
    required = ("web_app_url", "spreadsheet_id", "sheet_id", "token")
    missing = [key for key in required if not str(config.get(key, "")).strip()]
    if missing:
        raise ValueError("Google Sheets is not configured: " + ", ".join(missing))

    path = Path(workbook_path)
    if not path.exists():
        raise FileNotFoundError(f"Generated workbook not found: {path}")

    wb = load_workbook(path, data_only=False, read_only=False)
    try:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Worksheet {sheet_name!r} is not present in {path.name}")
        ws = wb[sheet_name]
        cells = []
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell) or cell.value is None:
                    continue
                cells.append({
                    "row": cell.row,
                    "col": cell.column,
                    "value": _serialise_cell(cell.value),
                })
        payload = {
            "action": "replaceSheetValues",
            "spreadsheetId": str(config["spreadsheet_id"]).strip(),
            "sheetId": int(str(config["sheet_id"]).strip()),
            "sheetName": sheet_name,
            "token": str(config["token"]),
            "rowCount": ws.max_row or 1,
            "columnCount": ws.max_column or 1,
            "cells": cells,
        }
    finally:
        wb.close()

    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        str(config["web_app_url"]).strip(),
        data=body,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            result = json.loads(response.read().decode("utf-8"))
    except Exception as exc:
        raise RuntimeError(
            "Could not reach the Google Sheets bridge. Verify the Apps Script "
            "deployment URL, access setting and network connection."
        ) from exc

    if not result.get("success"):
        raise RuntimeError(result.get("error") or "Google Sheets sync was rejected.")
    return {
        "sheet_name": sheet_name,
        "cells_written": result.get("cellsWritten", len(cells)),
        "message": result.get("message", "Google Sheet updated."),
    }
