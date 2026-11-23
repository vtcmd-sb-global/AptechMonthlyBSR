import unittest
from datetime import datetime
from openpyxl import Workbook

from import_historical import (
    _faculty_header_text,
    _parse_faculty_line,
    _resolve_faculty_initials,
    parse_sheet_allocations,
    sheet_to_year_month,
)


class BsrImportTests(unittest.TestCase):
    def setUp(self):
        self.maps = {
            "faculty": {"T": 10, "AA": 11, "HM": 12, "AW": 13},
            "faculty_names": {
                "talha qureshi": "T",
                "aman ali": "AA",
                "hafiz moiz": "HM",
                "abdul wahab": "AW",
            },
            "labs": {"L1": 1},
            "slots": {"B": 1, "C": 2, "D": 3},
            "days": {"MWF": 1, "TTS": 2},
            "careers": {"CPISM": 1},
        }

    def test_month_names_and_spelling_variant(self):
        self.assertEqual(sheet_to_year_month("November-26"), "2026-11")
        self.assertEqual(sheet_to_year_month("Feburary-26"), "2026-02")
        self.assertIsNone(sheet_to_year_month("Sheet14"))

    def test_full_faculty_display_names_resolve_to_roster_initials(self):
        self.assertEqual(_parse_faculty_line("(AI) Talha")[0], "TALHA")
        self.assertEqual(_resolve_faculty_initials("Talha", self.maps), "T")
        self.assertEqual(_resolve_faculty_initials("Moiz", self.maps), "HM")
        self.assertEqual(_resolve_faculty_initials("A.Wahab", self.maps), "AW")
        self.assertEqual(_resolve_faculty_initials("Amaan", self.maps), "AA")

    def test_split_faculty_header_is_reassembled_from_block_cells(self):
        grid = {(18, 30): "(AI)", (18, 33): "Aousaja"}
        combined = _faculty_header_text(grid, 18, 30)
        self.assertEqual(combined, "(AI) Aousaja")
        self.assertEqual(_parse_faculty_line(combined), ("AOUSAJA", "AI"))

    def test_adjacent_d_slot_headers_stay_inside_their_own_blocks(self):
        grid = {
            (4, 26): "(PMTZ) Amaan",
            (4, 30): "(CDMA) Amaan",
        }
        mwf_header = _faculty_header_text(grid, 4, 26)
        tts_header = _faculty_header_text(grid, 4, 30)

        self.assertEqual(mwf_header, "(PMTZ) Amaan")
        self.assertEqual(tts_header, "(CDMA) Amaan")
        self.assertEqual(_parse_faculty_line(mwf_header), ("AMAAN", "PMTZ"))
        self.assertEqual(_parse_faculty_line(tts_header), ("AMAAN", "CDMA"))

    def test_imports_adjacent_d_slot_allocations_without_crossing_headers(self):
        wb = Workbook()
        ws = wb.active
        ws["A6"] = "L1"

        # D-slot MWF occupies columns 26-29; D-slot TTS starts at 30.
        ws.cell(4, 26, "(PMTZ) Amaan")
        ws.cell(5, 26, "PMTZ-202405B")
        ws.cell(6, 26, "FBD")
        ws.cell(7, 26, datetime(2024, 5, 4))
        ws.cell(8, 26, "ADSE II")
        ws.cell(8, 29, 4)

        ws.cell(4, 30, "(CDMA) Amaan")
        ws.cell(5, 30, "DM-202609D")
        ws.cell(6, 30, "SEO")
        ws.cell(7, 30, datetime(2026, 9, 8))
        ws.cell(8, 30, "SEM I")
        ws.cell(8, 34, 3)

        rows = parse_sheet_allocations(ws, "2026-11", self.maps, None)
        d_rows = [x for x in rows if x["time_slot_id"] == 3]
        self.assertEqual(len(d_rows), 2)
        self.assertEqual({x["day_group_id"] for x in d_rows}, {1, 2})
        self.assertEqual({x["career_code"] for x in d_rows}, {"ADSE II", "SEM I"})
        self.assertEqual({x["students"] for x in d_rows}, {3, 4})

    def test_imports_faculty_allocation_and_admission_open_slot(self):
        wb = Workbook()
        ws = wb.active
        ws["A6"] = "L1"
        ws["B4"] = "(AI) Talha"
        ws["B5"] = "AI-202609B"
        ws["B6"] = "PWD"
        ws["B7"] = datetime(2026, 9, 19)
        ws["B8"] = "CPISM"
        ws["F8"] = 8
        ws["N4"] = "Admission Open"

        rows = parse_sheet_allocations(ws, "2026-11", self.maps, None)
        self.assertEqual(len(rows), 2)

        faculty_row = next(x for x in rows if x["faculty_id"] is not None)
        self.assertEqual(faculty_row["faculty_id"], 10)
        self.assertEqual(faculty_row["career_code"], "CPISM")
        self.assertEqual(faculty_row["students"], 8)

        open_row = next(x for x in rows if x["faculty_id"] is None)
        self.assertIsNone(open_row["career_id"])
        self.assertTrue(open_row["is_admission_open"])
        self.assertEqual(open_row["students"], 0)


if __name__ == "__main__":
    unittest.main()
