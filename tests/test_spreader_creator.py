import importlib.util
from pathlib import Path
import unittest


DEPENDENCIES_AVAILABLE = all(importlib.util.find_spec(name) for name in ("customtkinter", "openpyxl", "bs4"))
if DEPENDENCIES_AVAILABLE:
    from spreader_creator import calculate_times, parse_html_data, subtract_hours, validate_date


FIXTURE = Path(__file__).parent / "sample_worksheet.html"


@unittest.skipUnless(DEPENDENCIES_AVAILABLE, "application dependencies are not installed")
class SpreaderCreatorTests(unittest.TestCase):
    def test_fos_crosses_midnight(self):
        self.assertEqual(subtract_hours("0030", 2), "2230")

    def test_standard_shift(self):
        self.assertEqual(calculate_times("0800", "", ""), ("1600", ""))

    def test_date_validation(self):
        self.assertEqual(validate_date("2-03-2026"), "2-03-2026")
        with self.assertRaises(ValueError):
            validate_date("2026/02/03")

    def test_fixture_parser(self):
        data = parse_html_data(str(FIXTURE), lambda _: None)
        self.assertEqual(len(data.shifts), 2)
        self.assertEqual(data.leave_list, ["CD - Annual Leave (0800 - 1200)"])
        self.assertEqual(data.notes_list, ["Training assignment"])


if __name__ == "__main__":
    unittest.main()
