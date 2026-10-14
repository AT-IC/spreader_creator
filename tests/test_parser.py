from pathlib import Path
import importlib.util
import unittest

FIXTURE = Path(__file__).parent / "fixtures" / "sample_worksheet.html"


@unittest.skipUnless(importlib.util.find_spec("bs4"), "beautifulsoup4 is not installed")
class ParserTests(unittest.TestCase):
    def test_sample_worksheet_is_parsed(self):
        from parser import parse_html_data

        data = parse_html_data(str(FIXTURE))
        self.assertEqual(len(data.shifts), 2)
        self.assertEqual(data.shifts[0]["initials"], "AB")
        self.assertEqual(data.shifts[0]["code"], "A")
        self.assertEqual(data.leave_list, ["CD - Annual Leave (0800 - 1200)"])
        self.assertEqual(data.notes_list, ["Training assignment"])


if __name__ == "__main__":
    unittest.main()
