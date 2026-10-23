"""Resolve configured AMFI codes against the official NAV master; fail on ambiguity."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def main() -> None:
    nav_path = ROOT / "data" / "raw" / "amfi" / "nav.txt"
    digest = hashlib.sha256(nav_path.read_bytes()).hexdigest()
    lines = nav_path.read_text(encoding="utf-8-sig").splitlines()
    registry = json.loads((ROOT / "config" / "schemes.json").read_text(encoding="utf-8"))
    mappings = []
    for scheme in registry["schemes"]:
        code = scheme["scheme_code"]
        hits = []
        for line_number, line in enumerate(lines, 1):
            cols = line.split(";")
            if len(cols) < 8 or cols[0].strip() != code:
                continue
            plan = cols[4].strip().casefold()
            option = cols[5].strip().casefold()
            if plan == "direct plan" and "growth" in option:
                hits.append((line_number, cols))
        if len(hits) != 1:
            raise SystemExit(f"AMFI code {code} resolved to {len(hits)} Direct Growth rows; halt")
        line_number, cols = hits[0]
        current_name = cols[3].strip()
        if not current_name.casefold().startswith(scheme["name"].casefold()):
            raise SystemExit(f"AMFI code {code} name mismatch: {current_name!r}; halt")
        mappings.append(
            {
                "scheme_code": code,
                "configured_name": scheme["name"],
                "amfi_name": current_name,
                "isin": cols[1].strip(),
                "plan": cols[4].strip(),
                "option": cols[5].strip(),
                "nav_effective_date": cols[7].strip(),
                "source_file": "data/raw/amfi/nav.txt",
                "source_line": line_number,
                "source_sha256": digest,
            }
        )
    output = {"registry_source_sha256": digest, "mappings": mappings}
    destination = ROOT / "data" / "scheme-resolution.json"
    destination.write_text(json.dumps(output, indent=2) + "\n", encoding="utf-8")
    print(f"Resolved {len(mappings)} scheme codes; all are unambiguous Direct Growth rows")


if __name__ == "__main__":
    main()
