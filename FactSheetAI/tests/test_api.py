import shutil
import unittest
from pathlib import Path

from fastapi import HTTPException

from factsheetai.api import (
    analyze,
    analyze_portfolio,
    catalog,
    export_excel,
    get_thresholds,
    health,
    upload_preview,
)


class ApiTests(unittest.TestCase):
    def test_health_and_dynamic_catalog(self):
        self.assertEqual(health(), {"status": "ok"})
        result = catalog()
        available = {row["scheme_code"] for row in result["schemes"] if row["available_months"]}
        self.assertTrue({"119018", "118632", "119160", "119063"}.issubset(available))
        self.assertEqual(result["months"][0], "2025-07")
        self.assertEqual(result["months"][-1], "2026-06")

    def test_source_backed_analysis_and_validation(self):
        data = analyze("2025-07", "119018", "118632")
        self.assertTrue(data["validation"]["match"])
        self.assertEqual(data["provenance"]["amfi_top100"]["effective_date"], "2025-06-30")
        self.assertTrue(data["common_holdings"])
        self.assertEqual(len(data["provenance"]["a"]["sha256"]), 64)

    def test_absent_primary_scheme_is_reported_unavailable(self):
        with self.assertRaises(HTTPException) as caught:
            analyze("2025-07", "120586", "119018")
        self.assertEqual(caught.exception.status_code, 404)
        self.assertIn("Data not available", caught.exception.detail)

    def test_out_of_window_month_is_rejected(self):
        with self.assertRaises(HTTPException) as caught:
            analyze("2026-13", "119018", "119063")
        self.assertEqual(caught.exception.status_code, 422)

    def test_multi_fund_endpoint_returns_matrix_flags_and_validation(self):
        data = analyze_portfolio("2025-07", "119018,118632")
        self.assertEqual(set(data["overlap_matrix"]), {"119018", "118632"})
        self.assertEqual(len(data["pairs"]), 1)
        self.assertIn("pair_recomputations", data["validation"])
        self.assertIn("health_scores", data)
        self.assertIsNone(data["active_share"])

    def test_threshold_config_is_exposed(self):
        self.assertIn("HIGH_PAIR_OVERLAP", get_thresholds()["rules"])

    def test_normalized_csv_upload_is_previewed_not_imported(self):
        csv = (Path(__file__).parent / "fixtures" / "SYNTHETIC_a.csv").read_text(encoding="utf-8")
        preview = upload_preview({"csv": csv})
        self.assertEqual(preview["status"], "REVIEW_REQUIRED")
        self.assertEqual(preview["holdings_count"], 3)
        self.assertFalse(preview["weights_near_100"])

    @unittest.skipIf(shutil.which("libreoffice") or shutil.which("soffice"), "LibreOffice is installed")
    def test_excel_export_fails_closed_without_recalculation_runtime(self):
        with self.assertRaises(HTTPException) as caught:
            export_excel({})
        self.assertEqual(caught.exception.status_code, 503)


if __name__ == "__main__":
    unittest.main()
