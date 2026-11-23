import json
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from google_sheets_sync import (
    get_google_sheets_sync_state,
    mark_google_sheets_month_synced,
    queue_google_sheets_sync,
    record_google_sheets_sync_error,
    save_google_sheets_config,
    sync_workbook_to_google_sheet,
)


class _FakeResponse:
    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def read(self):
        return b'{"success": true, "cellsWritten": 2, "message": "ok"}'


class GoogleSheetsSyncTests(unittest.TestCase):
    def test_payload_targets_month_tab_and_skips_merged_placeholders(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "bsr.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "November-26"
            ws.merge_cells("A1:B1")
            ws["A1"] = "GLS Labstatus"
            ws["C2"] = datetime(2026, 11, 1)
            wb.save(path)
            wb.close()

            config = {
                "web_app_url": "https://example.invalid/exec",
                "spreadsheet_id": "spreadsheet-id",
                "sheet_id": "1463201093",
                "token": "test-secret",
            }
            with patch("google_sheets_sync.urllib.request.urlopen", return_value=_FakeResponse()) as urlopen:
                result = sync_workbook_to_google_sheet(path, "November-26", config)

            request = urlopen.call_args.args[0]
            payload = json.loads(request.data.decode("utf-8"))

            self.assertEqual(payload["sheetName"], "November-26")
            self.assertEqual(payload["spreadsheetId"], "spreadsheet-id")
            self.assertEqual(payload["sheetId"], 1463201093)
            self.assertEqual(payload["rowCount"], 2)
            self.assertEqual(payload["columnCount"], 3)
            self.assertEqual(len(payload["cells"]), 2)
            self.assertEqual(payload["cells"][1]["value"]["type"], "date")
            self.assertEqual(result["cells_written"], 2)


    def test_sync_queue_is_durable_and_clears_month_after_success(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            config = {
                "web_app_url": "https://example.invalid/exec",
                "spreadsheet_id": "spreadsheet-id",
                "sheet_id": "1463201093",
                "token": "test-secret",
            }
            with patch("google_sheets_sync.CONFIG_PATH", root / "config.json"), patch(
                "google_sheets_sync.SYNC_STATE_PATH", root / "sync-state.json"
            ):
                save_google_sheets_config(config)
                pending = queue_google_sheets_sync(["2026-11", "invalid", "2026-10"])
                self.assertEqual(pending, ["2026-10", "2026-11"])

                record_google_sheets_sync_error("Network unavailable")
                self.assertEqual(
                    get_google_sheets_sync_state()["pending_months"],
                    ["2026-10", "2026-11"],
                )
                self.assertEqual(
                    get_google_sheets_sync_state()["last_error"],
                    "Network unavailable",
                )

                mark_google_sheets_month_synced("2026-10")
                state = get_google_sheets_sync_state()
                self.assertEqual(state["pending_months"], ["2026-11"])
                self.assertEqual(state["last_success_month"], "2026-10")
                self.assertIsNotNone(state["last_success_at"])

    def test_sync_queue_does_not_queue_before_configuration(self):
        with tempfile.TemporaryDirectory() as tmp:
            with patch("google_sheets_sync.CONFIG_PATH", Path(tmp) / "missing.json"), patch(
                "google_sheets_sync.SYNC_STATE_PATH", Path(tmp) / "sync-state.json"
            ):
                self.assertEqual(queue_google_sheets_sync(["2026-11"]), [])
                self.assertFalse((Path(tmp) / "sync-state.json").exists())


if __name__ == "__main__":
    unittest.main()
