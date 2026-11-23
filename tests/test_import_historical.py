import unittest
from datetime import datetime
from openpyxl import Workbook

from import_historical import (
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
            "slots": {"B": 1, "C": 2},
            "days": {"MWF": 1},
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
