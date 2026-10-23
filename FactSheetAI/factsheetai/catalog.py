"""Versioned scheme registry and dynamic access to authentic monthly sources."""

import json
import logging
from pathlib import Path
from typing import Any

from .ingestion import load_csv, load_hdfc_xlsx, load_nippon_xls, load_sbi_xlsx, load_tata_xlsx
from .models import PortfolioSnapshot

ROOT = Path(__file__).resolve().parents[1]
REGISTRY = json.loads((ROOT / "config" / "schemes.json").read_text(encoding="utf-8"))
SCHEMES: dict[str, dict[str, Any]] = {row["scheme_code"]: row for row in REGISTRY["schemes"]}
LOGGER = logging.getLogger(__name__)
SNAPSHOT_CACHE: dict[tuple[str, str, str, int], PortfolioSnapshot] = {}


def _csv_snapshots(raw_dir: Path) -> list[tuple[Path, PortfolioSnapshot]]:
    """Load normalized source files; this root is never pointed at test fixtures by the app."""
    found = []
    for path in raw_dir.rglob("*.csv"):
        try:
            found.append((path, load_csv(path)))
        except (ValueError, OSError) as exc:
            LOGGER.warning("Skipping invalid normalized source %s: %s", path, exc)
    return found


def load_snapshot(
    scheme_code: str, month: str, *, raw_dir: Path | None = None
) -> PortfolioSnapshot:
    """Load an official adapter file or a normalized, user-supplied CSV snapshot."""
    raw_dir = raw_dir or ROOT / "data" / "raw"
    scheme = SCHEMES.get(scheme_code)
    adapter = scheme.get("adapter") if scheme else None
    source: Path | None = None
    parser = None
    if adapter == "hdfc_xlsx":
        stem = f"nifty50-{month}" if scheme_code == "119063" else month
        source, parser = raw_dir / "amc" / "hdfc" / f"{stem}.xlsx", load_hdfc_xlsx
    elif adapter == "nippon_xls":
        source, parser = raw_dir / "amc" / "nippon-india" / f"{month}.xls", load_nippon_xls
    elif adapter == "tata_xlsx":
        source, parser = raw_dir / "amc" / "tata" / f"{month}.xlsx", load_tata_xlsx
    elif adapter == "sbi_xlsx":
        source, parser = raw_dir / "amc" / "sbi" / f"{month}.xlsx", load_sbi_xlsx
    if source is not None and source.is_file() and parser is not None:
        if scheme is None:
            raise FileNotFoundError(f"No scheme metadata configured for {scheme_code}")
        stat = source.stat()
        cache_key = (scheme_code, month, str(source.resolve()), stat.st_mtime_ns)
        cached = SNAPSHOT_CACHE.get(cache_key)
        if cached is not None:
            return cached.model_copy(deep=True)
        snapshot = parser(source, scheme_id=scheme_code, scheme_name=scheme["name"])
        if snapshot.disclosure_month.strftime("%Y-%m") != month:
            raise ValueError(f"Source workbook date does not match requested month {month}")
        SNAPSHOT_CACHE[cache_key] = snapshot
        return snapshot.model_copy(deep=True)

    matches = [
        snapshot
        for _, snapshot in _csv_snapshots(raw_dir)
        if snapshot.scheme_id == scheme_code
        and snapshot.disclosure_month.strftime("%Y-%m") == month
    ]
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise ValueError(f"multiple normalized files found for {scheme_code} {month}")
    if scheme is None:
        raise FileNotFoundError(f"Data not available for scheme code {scheme_code} in {month}")
    if source is None:
        source = raw_dir / "schemes" / scheme_code / f"{month}.csv"
    raise FileNotFoundError(f"Missing source file: {source}")


def available_months(scheme_code: str, *, raw_dir: Path | None = None) -> list[str]:
    raw_dir = raw_dir or ROOT / "data" / "raw"
    scheme = SCHEMES.get(scheme_code)
    adapter = scheme.get("adapter") if scheme else None
    months: set[str] = {
        snapshot.disclosure_month.strftime("%Y-%m")
        for _, snapshot in _csv_snapshots(raw_dir)
        if snapshot.scheme_id == scheme_code
    }
    if adapter == "hdfc_xlsx":
        prefix = "nifty50-" if scheme_code == "119063" else ""
        folder = raw_dir / "amc" / "hdfc"
        months.update(p.stem.removeprefix(prefix) for p in folder.glob(f"{prefix}20??-??.xlsx"))
    elif adapter == "nippon_xls":
        folder = raw_dir / "amc" / "nippon-india"
        months.update(p.stem for p in folder.glob("20??-??.xls"))
    elif adapter == "tata_xlsx":
        folder = raw_dir / "amc" / "tata"
        months.update(p.stem for p in folder.glob("20??-??.xlsx"))
    elif adapter == "sbi_xlsx":
        folder = raw_dir / "amc" / "sbi"
        months.update(p.stem for p in folder.glob("20??-??.xlsx"))
    return sorted(months)


def catalog_response(*, raw_dir: Path | None = None) -> dict[str, object]:
    raw_dir = raw_dir or ROOT / "data" / "raw"
    entries: list[dict[str, Any]] = []
    for code, scheme in SCHEMES.items():
        months = available_months(code, raw_dir=raw_dir)
        entries.append(
            {
                "scheme_code": code,
                "name": scheme["name"],
                "role": scheme["role"],
                "available_months": months,
                "status": "available" if months else "Data not available",
            }
        )
    by_code = {str(entry["scheme_code"]): entry for entry in entries}
    for _, snapshot in _csv_snapshots(raw_dir):
        entry = by_code.get(snapshot.scheme_id)
        if entry is None:
            entry = {
                "scheme_code": snapshot.scheme_id,
                "name": snapshot.scheme_name,
                "role": "user_supplied",
                "available_months": [],
                "status": "available",
            }
            by_code[snapshot.scheme_id] = entry
            entries.append(entry)
        listed_months: list[str] = entry["available_months"]
        month = snapshot.disclosure_month.strftime("%Y-%m")
        if month not in listed_months:
            listed_months.append(month)
            listed_months.sort()
            entry["status"] = "available"
    return {
        "schemes": entries,
        "months": sorted({month for entry in entries for month in entry["available_months"]}),
    }
