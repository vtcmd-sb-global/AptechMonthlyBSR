import unittest

from export_excel import _block_end_col, _student_value_col


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


if __name__ == "__main__":
    unittest.main()
