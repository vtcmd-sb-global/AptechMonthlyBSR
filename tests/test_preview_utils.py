import unittest

from ui.preview_utils import preferred_sheet_name


class PreviewSheetSelectionTests(unittest.TestCase):
    def setUp(self):
        self.names = [
            "May-25", "August-25", "February-26", "August-26",
            "November-26", "Chart1", "Dashboard Summary", "_Data",
        ]

    def test_generated_report_opens_its_own_month(self):
        self.assertEqual(
            preferred_sheet_name(self.names, "Output/BSR_2026-08.xlsx"),
            "August-26",
        )
        self.assertEqual(
            preferred_sheet_name(self.names, "Output/BSR_2026-11.xlsx"),
            "November-26",
        )

    def test_nonstandard_filename_falls_back_to_newest_month(self):
        self.assertEqual(
            preferred_sheet_name(self.names, "uploaded-workbook.xlsx"),
            "November-26",
        )

    def test_month_name_spelling_variant_is_supported(self):
        names = ["Feburary-26", "January-26", "Chart1"]
        self.assertEqual(
            preferred_sheet_name(names, "uploaded-workbook.xlsx"),
            "Feburary-26",
        )

    def test_abbreviated_and_iso_month_tabs_are_supported(self):
        names = ["Aug-26", "2026-09", "Nov-26", "Chart1"]
        self.assertEqual(
            preferred_sheet_name(names, "uploaded-workbook.xlsx"),
            "Nov-26",
        )

    def test_helper_tabs_are_not_selected_when_no_month_exists(self):
        names = ["Dashboard Summary", "_Data", "Chart1", "Allocations", "BSR", "Overview"]
        self.assertEqual(
            preferred_sheet_name(names, "uploaded-workbook.xlsx"),
            "Overview",
        )


if __name__ == "__main__":
    unittest.main()
