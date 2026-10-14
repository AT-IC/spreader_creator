import tempfile
import unittest
from pathlib import Path

from excel_writer import build_excel_file
from models import ParsedData


class ExcelWriterTests(unittest.TestCase):
    def test_dry_run_does_not_require_template_or_write_file(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            data = ParsedData([], [], [])
            result = build_excel_file(data, "2-03-2026", str(Path(temp_dir) / "missing.xlsm"), temp_dir, dry_run=True)
            self.assertIsNone(result)
            self.assertEqual(list(Path(temp_dir).iterdir()), [])


if __name__ == "__main__":
    unittest.main()
