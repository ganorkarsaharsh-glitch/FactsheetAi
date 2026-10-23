"""List every unavailable configured large-cap scheme-month disclosure."""

from __future__ import annotations

import json

from factsheetai.catalog import ROOT, SCHEMES, available_months

MONTHS = [f"2025-{month:02d}" for month in range(7, 13)] + [
    f"2026-{month:02d}" for month in range(1, 7)
]


def main() -> None:
    missing = []
    for code, scheme in SCHEMES.items():
        if scheme["role"] == "benchmark_control":
            continue
        available = set(available_months(code))
        for month in MONTHS:
            if month not in available:
                if code == "120586":
                    reason = (
                        "An official monthly factsheet PDF is archived, but it omits holding ISINs "
                        "and aggregates sub-1% equity positions; no complete ISIN-level portfolio "
                        "source is available for this month."
                    )
                    expected_file = f"data/raw/amc/icici/{month}.<ISIN-bearing portfolio file>"
                else:
                    reason = "No official monthly portfolio file for this scheme-month is present in data/raw."
                    expected_file = f"data/raw/amc/<adapter>/{month}.<source-extension>"
                missing.append(
                    {
                        "scheme_code": code,
                        "scheme_name": scheme["name"],
                        "disclosure_month": month,
                        "expected_file": expected_file,
                        "reason": reason,
                    }
                )
    output = {
        "window": [MONTHS[0], MONTHS[-1]],
        "required_months": len(MONTHS),
        "missing_scheme_month_count": len(missing),
        "missing_scheme_months": missing,
        "note": "No holdings, identifiers or values were inferred for missing files.",
    }
    path = ROOT / "data" / "missing-source-files.json"
    path.write_text(json.dumps(output, indent=2) + "\n", encoding="utf-8")
    print(f"Missing source report written: {len(missing)} scheme-month files")


if __name__ == "__main__":
    main()
