"""Rebuild the immutable raw-file hash manifest from recorded source URLs."""

from __future__ import annotations

import hashlib
import json
from datetime import UTC, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def main() -> None:
    sources = json.loads((RAW / "source_index.json").read_text(encoding="utf-8"))
    entries = []
    for source in sources:
        path = ROOT / source["file"]
        if not path.is_file():
            continue
        entries.append(
            {
                "file": source["file"].replace("\\", "/"),
                "amc": source["amc"],
                "disclosure_month": source.get("disclosure_month"),
                "effective_date": source.get("effective_date"),
                "source_type": source.get("source_type", "source_data"),
                "engine_eligible": source.get("engine_eligible", True),
                "source_url": source["url"],
                "bytes": path.stat().st_size,
                "sha256": sha256(path),
            }
        )
    entries.sort(key=lambda row: row["file"])
    output = {
        "manifest_version": 1,
        "generated_at": datetime.now(UTC).isoformat(),
        "source_count": len(entries),
        "total_bytes": sum(row["bytes"] for row in entries),
        "files": entries,
    }
    (RAW / "manifest.json").write_text(json.dumps(output, indent=2) + "\n", encoding="utf-8")
    print(f"Manifest written: {len(entries)} files, {output['total_bytes']:,} bytes")


if __name__ == "__main__":
    main()
