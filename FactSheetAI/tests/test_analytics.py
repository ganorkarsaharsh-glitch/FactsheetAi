import unittest
from decimal import Decimal as D
from pathlib import Path

from hypothesis import given
from hypothesis import strategies as st

from factsheetai.analytics import (
    active_share,
    assess_rule,
    combined_portfolio_detail,
    holding_drift,
    identity_overlap,
    marginal_overlap,
    overlap,
    overlap_matrix,
    overlap_rank,
    pandas_overlap,
    portfolio_concentration,
    python_overlap,
    red_flags,
    style_drift,
    top_holdings_stats,
    turnover,
    weighted_portfolio,
)
from factsheetai.ingestion import load_csv

FIXTURES = Path(__file__).parent / "fixtures"


class AnalyticsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.a = load_csv(FIXTURES / "SYNTHETIC_a.csv")
        cls.b = load_csv(FIXTURES / "SYNTHETIC_b.csv")
        cls.style = load_csv(FIXTURES / "SYNTHETIC_style.csv")
        cls.concentration = load_csv(FIXTURES / "SYNTHETIC_concentration.csv")

    def test_overlap_and_independent_paths(self):
        result = overlap(self.a, self.b)
        self.assertEqual(result["overlap_pct"], D(5))
        self.assertEqual(result["jaccard"], D(1) / D(3))
        self.assertEqual(pandas_overlap(self.a, self.b), python_overlap(self.a, self.b))
        self.assertEqual(result["common_isins"], ["INE000A01001"])

    def test_matrix_symmetry_and_diagonal(self):
        matrix = overlap_matrix({"a": self.a, "b": self.b})
        self.assertEqual(matrix["a"]["b"], matrix["b"]["a"])
        self.assertEqual(matrix["a"]["a"], D(10))

    def test_concentration_hhi_and_effective_stocks(self):
        stats = top_holdings_stats(self.concentration, 2)
        self.assertEqual(stats["top_n_concentration_pct"], D(5))
        self.assertEqual(stats["equity_hhi"], D("0.0014"))
        self.assertEqual(stats["effective_stocks"], D(1) / D("0.0014"))

    def test_style_drift_uses_equities_only_and_honours_floor(self):
        result = style_drift(self.style, {"INE000A01001"})
        self.assertEqual(result["top100_weight_pct"], D(80))
        self.assertFalse(result["mandate_breach"])
        self.assertTrue(style_drift(self.style, set())["mandate_breach"])

    def test_weighted_portfolio_and_marginal_overlap(self):
        composite = weighted_portfolio({"a": self.a, "b": self.b}, {"a": D(50), "b": D(50)})
        self.assertEqual(composite["INE000A01001"], D("6.5"))
        concentration = portfolio_concentration(
            {"a": self.a, "b": self.b}, {"a": D(50), "b": D(50)}
        )
        self.assertEqual(concentration["unique_stocks"], D(3))
        self.assertGreater(concentration["effective_stocks"], D(0))
        self.assertEqual(marginal_overlap({"a": self.a}, {"a": D(100)}, self.b), D(5))

    def test_turnover_and_active_share(self):
        self.assertEqual(turnover(self.a, self.b), D("4.5"))
        self.assertEqual(active_share(self.a, self.b), D(7))
        self.assertEqual(identity_overlap(self.a, self.b), D(5))

    def test_combined_portfolio_and_monthly_drift_hand_examples(self):
        detail = combined_portfolio_detail({"a": self.a, "b": self.b}, {"a": D(100), "b": D(100)})
        self.assertEqual(detail["unique_stocks"], 3)
        self.assertEqual(detail["equity_rupees"], D(19))
        removed = combined_portfolio_detail({"a": self.a, "b": self.b}, {"a": D(100), "b": D(0)})
        self.assertEqual(removed["unique_stocks"], 2)
        self.assertEqual(removed["stock_positions"], 2)
        drift = holding_drift(self.a, self.b)
        self.assertEqual(drift["portfolio_change_pct"], D("4.5"))
        self.assertEqual(len(drift["new_entries"]), 1)
        self.assertEqual(len(drift["exits"]), 1)

    def test_rule_thresholds_apply_amber_margin_and_below_direction(self):
        rule = {"threshold": 50, "direction": "above", "amber_margin_pct": 15}
        self.assertEqual(assess_rule(rule, D(50)), "AMBER")
        self.assertEqual(assess_rule(rule, D(43)), "AMBER")
        self.assertEqual(assess_rule(rule, D(42)), "OK")
        self.assertEqual(assess_rule({"threshold": 80, "direction": "below"}, D(79)), "RED")
        self.assertEqual(assess_rule({"threshold": 50, "direction": "above", "inclusive": True}, D(50)), "RED")
        self.assertEqual([overlap_rank(D(x)) for x in (50, 30, 29)], ["HIGH", "MODERATE", "LOW"])

    def test_red_flags_are_deterministic(self):
        alerts = red_flags(self.style, top100_isins=set(), max_top10_pct=D(100))
        self.assertEqual([item["code"] for item in alerts], ["LARGE_CAP_FLOOR"])

    @given(st.integers(min_value=1, max_value=100), st.integers(min_value=1, max_value=100))
    def test_overlap_is_symmetric_and_bounded(self, left, right):
        # The property varies fixture-backed weights without exporting generated data.

        a, b = self.a.model_copy(deep=True), self.b.model_copy(deep=True)
        a.holdings[0] = a.holdings[0].model_copy(update={"weight_pct": D(left)})
        b.holdings[0] = b.holdings[0].model_copy(update={"weight_pct": D(right)})
        ab, ba = overlap(a, b)["overlap_pct"], overlap(b, a)["overlap_pct"]
        self.assertEqual(ab, ba)
        self.assertLessEqual(ab, min(D(left), D(right)))


if __name__ == "__main__":
    unittest.main()
