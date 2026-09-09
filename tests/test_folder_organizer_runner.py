import tempfile
import unittest
from pathlib import Path

from copilots_app.services.folder_organizer.runner import FolderOrganizerRunner


class FolderOrganizerRunnerTests(unittest.TestCase):
    def test_parse_plan_accepts_valid_and_reports_invalid(self):
        runner = FolderOrganizerRunner()
        plan = """
# comment
MOVE "a/file.txt" -> "docs/file.txt"
INVALID LINE
MOVE src/main.py -> app/main.py
"""
        res = runner.parse_plan(plan)
        self.assertEqual(res["instruction_count"], 2)
        self.assertEqual(len(res["errors"]), 1)

    def test_execute_plan_copies_files_without_touching_source(self):
        with tempfile.TemporaryDirectory() as tmp:
            src_root = Path(tmp) / "src"
            src_root.mkdir(parents=True, exist_ok=True)
            (src_root / "alpha.txt").write_text("alpha", encoding="utf-8")

            runner = FolderOrganizerRunner()
            runner.set_source_root(str(src_root))
            res = runner.execute_plan('MOVE "alpha.txt" -> "grouped/alpha.txt"')

            self.assertTrue(res["success"])
            self.assertEqual(res["copied_count"], 1)
            self.assertTrue((src_root / "alpha.txt").exists())
            self.assertTrue((Path(res["output_root"]) / "grouped" / "alpha.txt").exists())


if __name__ == "__main__":
    unittest.main()

