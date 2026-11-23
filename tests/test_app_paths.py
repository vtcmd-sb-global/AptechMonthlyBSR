import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import app_paths


class AppPathsTests(unittest.TestCase):
    def test_resource_path_is_relative_to_bundled_resource_directory(self):
        with patch.object(app_paths, "RESOURCE_DIR", Path("/bundle")):
            self.assertEqual(
                app_paths.resource_path("templates", "master.xlsx"),
                Path("/bundle/templates/master.xlsx"),
            )

    def test_legacy_user_files_are_copied_without_overwriting_new_data(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            legacy = root / "legacy"
            user_data = root / "new-user-data"
            output = root / "new-output"
            old_template = legacy / "user_data" / "latest_imported_bsr.xlsx"
            old_template.parent.mkdir(parents=True)
            old_template.write_bytes(b"legacy-template")
            old_report = legacy / "Output" / "BSR_2026-11.xlsx"
            old_report.parent.mkdir(parents=True)
            old_report.write_bytes(b"legacy-report")

            with (
                patch.object(app_paths, "legacy_roots", return_value=(legacy,)),
                patch.object(app_paths, "USER_DATA_DIR", user_data),
                patch.object(app_paths, "OUTPUT_DIR", output),
            ):
                app_paths.migrate_legacy_user_files()
                self.assertEqual(
                    (user_data / "latest_imported_bsr.xlsx").read_bytes(),
                    b"legacy-template",
                )
                self.assertEqual(
                    (output / "BSR_2026-11.xlsx").read_bytes(),
                    b"legacy-report",
                )

                (user_data / "latest_imported_bsr.xlsx").write_bytes(b"new-template")
                (output / "BSR_2026-11.xlsx").write_bytes(b"new-report")
                app_paths.migrate_legacy_user_files()

                self.assertEqual(
                    (user_data / "latest_imported_bsr.xlsx").read_bytes(),
                    b"new-template",
                )
                self.assertEqual(
                    (output / "BSR_2026-11.xlsx").read_bytes(),
                    b"new-report",
                )


if __name__ == "__main__":
    unittest.main()
