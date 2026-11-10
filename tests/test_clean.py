import sys
import unittest
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.build_report import clean_frame, parse_dates, parse_money  # noqa: E402


class ParseMoneyTest(unittest.TestCase):
    def test_space_and_comma(self):
        self.assertEqual(parse_money("1 284,50"), 1284.5)

    def test_currency_suffix(self):
        self.assertEqual(parse_money("1 391,20 руб."), 1391.2)

    def test_ruble_sign(self):
        self.assertEqual(parse_money("968 400 ₽"), 968400.0)

    def test_plain_number(self):
        self.assertEqual(parse_money(1284.5), 1284.5)

    def test_empty(self):
        self.assertIsNone(parse_money(""))
        self.assertIsNone(parse_money(None))
        self.assertIsNone(parse_money("—"))


class ParseDatesTest(unittest.TestCase):
    def test_mixed_formats(self):
        parsed = parse_dates(pd.Series(["01.06.2024", "2024-07-18"]))
        self.assertEqual(parsed.iloc[0].strftime("%Y-%m-%d"), "2024-06-01")
        self.assertEqual(parsed.iloc[1].strftime("%Y-%m-%d"), "2024-07-18")

    def test_total_row_becomes_nat(self):
        parsed = parse_dates(pd.Series(["01.06.2024", "Итого"]))
        self.assertTrue(pd.isna(parsed.iloc[1]))


class CleanFrameTest(unittest.TestCase):
    def setUp(self):
        self.frame = pd.DataFrame(
            {
                "date": ["01.06.2024", "01.06.2024", "02.06.2024", "Итого"],
                "category": ["одежда ", "одежда ", " Обувь", ""],
                "manager": ["Иванов", "Иванов", "Петрова", ""],
                "qty": ["1", "1", "2", ""],
                "amount": ["1 284,50", "1 284,50", "968 400 ₽", "1 049 788,60"],
            }
        )

    def test_drops_total_rows(self):
        cleaned = clean_frame(self.frame, "test.csv")
        self.assertEqual(len(cleaned), 2)

    def test_removes_duplicates(self):
        cleaned = clean_frame(self.frame, "test.csv")
        self.assertEqual(len(cleaned), 2)

    def test_normalizes_category(self):
        cleaned = clean_frame(self.frame, "test.csv")
        self.assertEqual(sorted(cleaned["category"]), ["Обувь", "Одежда"])

    def test_sums_amounts(self):
        cleaned = clean_frame(self.frame, "test.csv")
        self.assertAlmostEqual(float(cleaned["amount"].sum()), 969684.5, places=2)


if __name__ == "__main__":
    unittest.main()
