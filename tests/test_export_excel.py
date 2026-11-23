import unittest
from types import SimpleNamespace

from openpyxl import Workbook

from export_excel import (
    _block_end_col,
    _student_value_col,
    _fill_lab_data,
    _fill_slot_summaries,
)


class BsrExportBlockTests(unittest.TestCase):
    def test_block_boundaries_match_the_template_grid(self):
        self.assertEqual(_block_end_col(26), 29)  # D-slot MWF is four columns
        self.assertEqual(_block_end_col(30), 35)  # D-slot TTS is six columns
        self.assertEqual(_block_end_col(48), 55)  # F-slot MWF is eight columns
        self.assertEqual(_block_end_col(68), 73)  # final G-slot TTS block

    def test_student_count_column_stays_inside_its_block(self):
        self.assertEqual(_student_value_col(26), 29)
        self.assertEqual(_student_value_col(30), 34)
        self.assertEqual(_student_value_col(48), 52)

    def test_d_slot_export_keeps_mwf_students_and_tts_career_separate(self):
        ws = Workbook().active
        mwf = SimpleNamespace(
            lab_code="L1", time_slot_label="D (1:00-3:00)", day_group_code="MWF",
            course_title="PMTZ", faculty_initials="AA", batch_code="PMTZ-202405B",
            module_name="FBD", module_start_date=None, actual_start_date=None,
            notes=None, career_code="ADSE II", students=4, is_admission_open=False,
        )
        tts = SimpleNamespace(
            lab_code="L1", time_slot_label="D (1:00-3:00)", day_group_code="TTS",
            course_title="CDMA", faculty_initials="AA", batch_code="DM-202609D",
            module_name="SEO", module_start_date=None, actual_start_date=None,
            notes=None, career_code="SEM I", students=3, is_admission_open=False,
        )

        _fill_lab_data(ws, [mwf, tts], {})
        _fill_slot_summaries(ws, [mwf, tts])

        self.assertEqual(ws.cell(8, 26).value, "ADSE II")
        self.assertEqual(ws.cell(8, 29).value, 4)
        self.assertEqual(ws.cell(8, 30).value, "SEM I")
        self.assertEqual(ws.cell(8, 34).value, 3)
        self.assertEqual(ws.cell(55, 29).value, 4)
        self.assertEqual(ws.cell(55, 34).value, 3)


if __name__ == "__main__":
    unittest.main()
