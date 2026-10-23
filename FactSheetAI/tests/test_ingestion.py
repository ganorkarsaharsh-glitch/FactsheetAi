import unittest
from decimal import Decimal as D
from pathlib import Path
from tempfile import NamedTemporaryFile

from openpyxl import Workbook

from factsheetai.catalog import load_snapshot
from factsheetai.ingestion import load_csv, load_hdfc_xlsx, load_nippon_xls, load_sbi_xlsx

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = Path(__file__).parent / "fixtures"
MONTHS = [f"2025-{m:02d}" for m in range(7, 13)] + [f"2026-{m:02d}" for m in range(1, 7)]


class IngestionTests(unittest.TestCase):
    def test_hash_and_row_provenance(self):
        snapshot = load_csv(FIXTURES / "SYNTHETIC_a.csv")
        holding = snapshot.holdings[0]
        self.assertEqual(len(holding.provenance.sha256), 64)
        self.assertEqual(holding.provenance.source_row, 2)
        self.assertEqual(holding.weight_pct, D(8))
        self.assertIn("SYNTHETIC", holding.provenance.source_file)

    def test_mixed_month_rejected(self):
        with self.assertRaisesRegex(ValueError, "exactly one scheme-month"):
            load_csv(FIXTURES / "SYNTHETIC_mixed_month.csv")

    def test_invalid_headers_and_empty_csv_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "CSV requires columns"):
            load_csv(FIXTURES / "SYNTHETIC_bad_header.csv")
        with self.assertRaisesRegex(ValueError, "contains no records"):
            load_csv(FIXTURES / "SYNTHETIC_empty.csv")

    def test_actual_hdfc_archive_months_parse_and_preserve_sources(self):
        for month in MONTHS:
            with self.subTest(month=month):
                snapshot = load_snapshot("119018", month)
                self.assertEqual(snapshot.disclosure_month.strftime("%Y-%m"), month)
                self.assertGreater(len(snapshot.holdings), 35)
                self.assertGreater(sum(h.weight_pct for h in snapshot.holdings), D(80))
                self.assertEqual(len(snapshot.holdings[0].provenance.sha256), 64)
                self.assertTrue(snapshot.holdings[0].provenance.source_span)

    def test_actual_nippon_archive_months_parse(self):
        for month in MONTHS:
            with self.subTest(month=month):
                snapshot = load_snapshot("118632", month)
                self.assertEqual(snapshot.disclosure_month.strftime("%Y-%m"), month)
                self.assertGreater(len(snapshot.holdings), 50)
                self.assertGreater(sum(h.weight_pct for h in snapshot.holdings), D(80))

    def test_actual_tata_archive_months_parse(self):
        for month in MONTHS:
            with self.subTest(month=month):
                snapshot = load_snapshot("119160", month)
                self.assertEqual(snapshot.disclosure_month.strftime("%Y-%m"), month)
                self.assertGreater(len(snapshot.holdings), 30)

    def test_actual_sbi_archive_months_parse(self):
        for month in MONTHS:
            with self.subTest(month=month):
                snapshot = load_snapshot("119598", month)
                self.assertEqual(snapshot.disclosure_month.strftime("%Y-%m"), month)
                self.assertGreater(len(snapshot.holdings), 25)
                self.assertGreater(sum(h.weight_pct for h in snapshot.holdings), D(80))
                self.assertEqual(len(snapshot.holdings[0].provenance.sha256), 64)
                if month == "2026-04":
                    dlf = next(h for h in snapshot.holdings if h.isin == "INE271C01023")
                    self.assertIn(
                        "weight derived as market value / GRAND TOTAL (AUM) * 100",
                        dlf.provenance.source_span,
                    )

    def test_sbi_rejects_unresolved_special_weights_and_conflicting_duplicates(self):
        scenarios = [
            ("missing_aum", "#", 10.0, "# weight cannot be derived"),
            ("missing_value", "#", None, "has # weight and no market value"),
            ("bad_weight", "not-a-number", 10.0, "unsupported weight"),
            ("duplicate_conflict", "1.0", 10.0, "conflicting security names"),
        ]
        for scenario, weight, market_value, error in scenarios:
            with self.subTest(scenario=scenario):
                workbook = Workbook()
                index = workbook.active
                index.title = "Index"
                index.append(["FUND", "SBI Large Cap Fund"])
                sheet = workbook.create_sheet("FUND")
                sheet.append(["Portfolio Statement as on", "31-Mar-2026"])
                for _ in range(3):
                    sheet.append([])
                sheet.append(
                    [
                        "Name of the Instrument",
                        "ISIN",
                        "Industry",
                        "Market value (Rs. in Lakhs)",
                        "% to AUM",
                    ]
                )
                sheet.append(["Equity & Equity Related", None, None, None, None])
                sheet.append(["Example Ltd", "INE000A01011", "Industry", market_value, weight])
                if scenario == "duplicate_conflict":
                    sheet.append(["Different Ltd", "INE000A01011", "Industry", 5.0, 1.0])
                if scenario != "missing_aum":
                    sheet.append([None, None, "GRAND TOTAL (AUM)", 1000.0, 100.0])
                with NamedTemporaryFile(
                    suffix=f"_SYNTHETIC_{scenario}.xlsx", dir=FIXTURES, delete=False
                ) as temp:
                    path = Path(temp.name)
                try:
                    workbook.save(path)
                    with self.assertRaisesRegex(ValueError, error):
                        load_sbi_xlsx(path, scheme_id="119598", scheme_name="SBI Large Cap Fund")
                finally:
                    path.unlink(missing_ok=True)

    def test_tata_duplicate_isin_is_summed_and_both_rows_are_recorded(self):
        snapshot = load_snapshot("119160", "2026-02")
        ntpc = next(h for h in snapshot.holdings if h.isin == "INE733E01010")
        self.assertEqual(ntpc.weight_pct, D("1.79"))
        self.assertEqual(ntpc.provenance.source_row, 34)
        self.assertIn("34,58", ntpc.provenance.source_span)
        self.assertEqual(ntpc.confidence, D("0.95"))

    def test_malformed_configured_schemes_remain_unavailable(self):
        with self.assertRaisesRegex(FileNotFoundError, "Missing source file"):
            load_snapshot("120586", "2025-07")

    def test_hdfc_adapter_accepts_reference_snapshot(self):
        snapshot = load_hdfc_xlsx(
            ROOT / "data/raw/amc/hdfc/nifty50-2025-07.xlsx",
            scheme_id="119063",
            scheme_name="HDFC Nifty 50 Index Fund",
        )
        self.assertEqual(len(snapshot.holdings), 50)
        self.assertEqual(snapshot.disclosure_month.strftime("%Y-%m"), "2025-07")

    def test_nippon_scheme_ambiguity_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "not found in Nippon workbook index"):
            load_nippon_xls(
                ROOT / "data/raw/amc/nippon-india/2025-07.xls",
                scheme_id="000000",
                scheme_name="Imaginary Fund",
            )


if __name__ == "__main__":
    unittest.main()
