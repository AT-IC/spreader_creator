import unittest

from business_rules import calculate_times, subtract_hours
from excel_writer import validate_date


class BusinessRulesTests(unittest.TestCase):
    def test_fos_crosses_midnight(self):
        self.assertEqual(subtract_hours("0030", 2), "2230")

    def test_standard_shift(self):
        self.assertEqual(calculate_times("0800", "", ""), ("1600", ""))

    def test_extended_overtime(self):
        self.assertEqual(calculate_times("0800", "$$", "", ot_hours=4), ("1600", "2000"))

    def test_date_is_normalized(self):
        self.assertEqual(validate_date("2-03-2026"), "2-03-2026")

    def test_invalid_date_is_rejected(self):
        with self.assertRaises(ValueError):
            validate_date("2026/02/03")


if __name__ == "__main__":
    unittest.main()
