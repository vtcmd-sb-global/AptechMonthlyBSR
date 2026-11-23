import json
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from google_sheets_sync import sync_workbook_to_google_sheet


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


if __name__ == "__main__":
    unittest.main()
