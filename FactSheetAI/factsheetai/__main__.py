"""Command line entry point."""

import argparse
import json

from .analytics import top_holdings_stats
from .ingestion import load_csv


def main() -> None:
    parser = argparse.ArgumentParser(description="FactSheetAI source-backed portfolio tools")
    parser.add_argument("validate", nargs="?", help="validate a normalized CSV portfolio file")
    parser.add_argument("path", nargs="?")
    args = parser.parse_args()
    if args.validate:
        snap = load_csv(args.path or args.validate)
        stats = top_holdings_stats(snap)
        print(
            json.dumps(
                {
                    "scheme_id": snap.scheme_id,
                    "scheme_name": snap.scheme_name,
                    "month": str(snap.disclosure_month),
                    "rows": len(snap.holdings),
                    "sha256": snap.holdings[0].provenance.sha256,
                    "top_10_concentration_pct": str(stats["top_n_concentration_pct"]),
                },
                indent=2,
            )
        )
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
