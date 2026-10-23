import unittest
from pathlib import Path

from factsheetai.catalog import catalog_response, load_snapshot

FIXTURES = Path(__file__).parent / "fixtures"


class DynamicBehaviourTests(unittest.TestCase):
    def test_new_source_month_and_new_scheme_need_no_code_change(self):
        # Test data is synthetic and isolated in tests/fixtures; production scans only data/raw.
        first = load_snapshot("777001", "2025-01", raw_dir=FIXTURES)
        added_month = load_snapshot("777001", "2025-02", raw_dir=FIXTURES)
        added_scheme = load_snapshot("777002", "2025-02", raw_dir=FIXTURES)
        self.assertEqual(first.scheme_name, "Synthetic Dynamic Alpha")
        self.assertEqual(added_month.disclosure_month.strftime("%Y-%m"), "2025-02")
        self.assertEqual(added_scheme.scheme_name, "Synthetic Dynamic Beta")
        response = catalog_response(raw_dir=FIXTURES)
        entries = {row["scheme_code"]: row for row in response["schemes"]}
        self.assertEqual(entries["777002"]["available_months"], ["2025-02"])


if __name__ == "__main__":
    unittest.main()
