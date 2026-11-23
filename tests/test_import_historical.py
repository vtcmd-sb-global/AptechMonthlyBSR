import unittest
from datetime import datetime
from openpyxl import Workbook

from import_historical import (
    _faculty_header_text,
    _parse_faculty_line,
    _resolve_faculty_initials,
    _is_date_like,
    _fmt_date,
    parse_sheet_allocations,
    sheet_to_year_month,
)


class BsrImportTests(unittest.TestCase):
    def setUp(self):
        self.maps = {
            "faculty": {"T": 10, "AA": 11, "HM": 12, "AW": 13, "HA": 14},
            "faculty_names": {
                "talha qureshi": "T",
                "aman ali": "AA",
                "hafiz moiz": "HM",
                "abdul wahab": "AW",
                "hasnain": "HA",
                "hasnain arif": "HA",
            },
            "labs": {"L1": 1},
            "slots": {"B": 1, "C": 2, "D": 3},
            "days": {"MWF": 1, "TTS": 2},
            "careers": {"CPISM": 1, "ADSE II": 2, "SEM I": 3},
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
        self.assertEqual(_resolve_faculty_initials("Hasnain", self.maps), "HA")
        self.assertEqual(_resolve_faculty_initials("Hasnain Arif", self.maps), "HA")

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

    def test_import_scans_rows_beyond_old_fixed_limit_and_keeps_custom_values(self):
        import sqlite3
        from import_historical import ensure_faculty, load_grid

        wb = Workbook()
        ws = wb.active
        ws["A56"] = "L8"
        ws["B55"] = "(AI) New Person"
        ws["B56"] = "CUSTOM-202611A"
        ws["B57"] = "Custom Module"
        ws["B58"] = datetime(2026, 11, 3)
        ws["B59"] = "NEW PROGRAM"
        ws["F59"] = 7
        ws["B60"] = "SECOND PROGRAM"
        ws["F60"] = 4
        ws["B61"] = None
        ws["B62"] = None
        # A populated cell beyond the former 75-column import limit is retained.
        ws.cell(70, 82, "far-right data")

        conn = sqlite3.connect(":memory:")
        conn.execute(
            "CREATE TABLE faculty (id INTEGER PRIMARY KEY, initials TEXT UNIQUE, "
            "full_name TEXT, status TEXT, active INTEGER)"
        )
        conn.execute(
            "CREATE TABLE careers (id INTEGER PRIMARY KEY, code TEXT UNIQUE, name TEXT, "
            "is_stc INTEGER, is_sp INTEGER, sort_order INTEGER)"
        )
        maps = {
            **self.maps,
            "faculty": dict(self.maps["faculty"]),
            "faculty_names": dict(self.maps["faculty_names"]),
            "labs": {**self.maps["labs"], "L8": 8},
            "careers": dict(self.maps["careers"]),
        }
        grid = load_grid(ws)
        self.assertEqual(grid[(70, 82)], "far-right data")
        rows = parse_sheet_allocations(ws, "2026-11", maps, conn)

        self.assertEqual(len(rows), 2)
        self.assertEqual({row["career_code"] for row in rows}, {"NEW PROGRAM", "SECOND PROGRAM"})
        self.assertEqual({row["students"] for row in rows}, {4, 7})
        self.assertTrue(all(row["faculty_id"] is not None for row in rows))
        self.assertTrue(all(row["faculty_initials"] for row in rows))
        self.assertTrue(all(row["lab_id"] == 8 for row in rows))
        conn.close()

    def test_month_parser_accepts_abbreviated_and_iso_month_names(self):
        self.assertEqual(sheet_to_year_month("Aug-26"), "2026-08")
        self.assertEqual(sheet_to_year_month("2026-08"), "2026-08")
        self.assertEqual(sheet_to_year_month("Sept 2026 BSR"), "2026-09")

    def test_common_excel_date_strings_are_not_misread_as_careers(self):
        self.assertTrue(_is_date_like("19/09/2026"))
        self.assertTrue(_is_date_like("09/19/2026"))
        self.assertEqual(_fmt_date("19/09/2026"), "2026-09-19")


if __name__ == "__main__":
    unittest.main()
