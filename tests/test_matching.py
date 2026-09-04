import unittest
from decimal import Decimal

from matching import find_maximum_match, parse_amount, parse_amounts


class ParseAmountTests(unittest.TestCase):
    def test_accepts_currency_formatting(self):
        self.assertEqual(parse_amount("$1,234.50"), Decimal("1234.50"))

    def test_rejects_invalid_negative_and_fractional_cent_values(self):
        for value in ("", "nope", "-1.00", "1.001", "NaN", "Infinity"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                parse_amount(value)


class MaximumMatchTests(unittest.TestCase):
    def test_combines_entries_for_an_exact_match(self):
        result = find_maximum_match(parse_amounts(["10", "20"]), parse_amounts(["30"]))
        self.assertEqual(result.matched_total, Decimal("30.00"))
        self.assertEqual(result.matched_indices_a, (0, 1))
        self.assertEqual(result.matched_indices_b, (0,))

    def test_duplicate_amounts_remain_distinct_entries(self):
        result = find_maximum_match(parse_amounts(["10", "10"]), parse_amounts(["10"]))
        self.assertEqual(result.matched_total, Decimal("10.00"))
        self.assertEqual(len(result.matched_indices_a), 1)

    def test_no_match_returns_zero_and_full_unmatched_totals(self):
        result = find_maximum_match(parse_amounts(["10"]), parse_amounts(["3"]))
        self.assertEqual(result.matched_total, Decimal("0.00"))
        self.assertEqual(result.unmatched_total_a, Decimal("10.00"))
        self.assertEqual(result.unmatched_total_b, Decimal("3.00"))

    def test_decimal_addition_is_exact(self):
        result = find_maximum_match(parse_amounts(["0.10", "0.20"]), parse_amounts(["0.30"]))
        self.assertEqual(result.matched_total, Decimal("0.30"))

    def test_finds_the_maximum_not_the_first_common_total(self):
        result = find_maximum_match(parse_amounts(["5", "7"]), parse_amounts(["6", "7"]))
        self.assertEqual(result.matched_total, Decimal("7.00"))

    def test_supports_a_larger_set_within_the_documented_limit(self):
        values = [str(number) for number in range(1, 19)]
        result = find_maximum_match(parse_amounts(values), parse_amounts(values))
        self.assertEqual(result.matched_total, Decimal("171.00"))
        self.assertEqual(len(result.matched_indices_a), 18)


if __name__ == "__main__":
    unittest.main()

