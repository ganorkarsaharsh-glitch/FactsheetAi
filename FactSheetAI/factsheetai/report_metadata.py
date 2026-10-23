"""Extract report metadata that is explicitly present in AMC portfolio workbooks."""

from __future__ import annotations

import json
import re
from functools import lru_cache
from pathlib import Path
from typing import Any

import pandas as pd

from .catalog import ROOT, SCHEMES, load_snapshot
from .factsheet_lookup import lookup_official_factsheet, openai_lookup_enabled


def _number(value: Any) -> float | None:
    if isinstance(value, bool) or value is None or pd.isna(value):
        return None
    text = str(value).replace(",", "").replace("₹", "").strip()
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    return float(match.group()) if match else None


def _scheme_sheet(path: Path, scheme_name: str) -> tuple[pd.ExcelFile, str]:
    book = pd.ExcelFile(path)
    index_sheet = next((name for name in book.sheet_names if name.casefold() == "index"), None)
    if index_sheet is None:
        return book, book.sheet_names[0]

    index = pd.read_excel(path, sheet_name=index_sheet, header=None, dtype=object)
    expected = " ".join(scheme_name.upper().split())
    for _, row in index.iterrows():
        cells = [str(value).strip() for value in row if pd.notna(value)]
        for position, cell in enumerate(cells):
            normalized_cell = " ".join(cell.upper().split())
            if (normalized_cell == expected or normalized_cell.startswith(expected)) and position:
                candidate = cells[position - 1]
                if candidate in book.sheet_names:
                    return book, candidate
    book.close()
    raise ValueError(f"Could not locate {scheme_name} in {path.name}")


def _rows(path: Path, scheme_name: str) -> tuple[pd.DataFrame, str]:
    book, sheet_name = _scheme_sheet(path, scheme_name)
    try:
        return pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object), sheet_name
    finally:
        book.close()


def _ratio_from_rows(rows: list[list[Any]]) -> float | None:
    for row in rows:
        cells = [str(value).strip() for value in row if pd.notna(value)]
        text = " ".join(cells)
        if "turnover" not in text.casefold():
            continue
        percent = re.search(r"turnover[^\d-]*(-?\d+(?:\.\d+)?)\s*%", text, re.IGNORECASE)
        if percent:
            return float(percent.group(1))
        times = re.search(r"turnover[^\d-]*(-?\d+(?:\.\d+)?)\s*times?", text, re.IGNORECASE)
        if times:
            return float(times.group(1)) * 100
        if "portfolio turnover ratio" in text.casefold():
            numbers = [_number(value) for value in row]
            values = [value for value in numbers if value is not None]
            if values:
                # SBI states the turnover ratio as a fraction (0.48 = 48%).
                return values[-1] * 100 if abs(values[-1]) <= 2 else values[-1]
    return None


@lru_cache(maxsize=64)
def _auto_values(path: Path, scheme_name: str) -> dict[str, Any]:
    frame, sheet_name = _rows(path, scheme_name)
    rows = frame.values.tolist()
    metadata: dict[str, Any] = {
        "aum_cr": None,
        "expense": None,
        "cat_expense": None,
        "turnover": _ratio_from_rows(rows),
        "cash": None,
    }
    for row in rows:
        cells = [str(value).strip() for value in row if pd.notna(value)]
        text = " ".join(cells).casefold()
        values = [_number(value) for value in row]
        numbers = [value for value in values if value is not None]
        if not numbers:
            continue

        # AMC statements report these total asset values in ₹ lakhs.
        if "grand total (aum)" in text or text.startswith("net assets") or text == "grand total":
            metadata["aum_cr"] = numbers[0] / 100

        if "cash, cash equivalents and net current assets" in text or "cash / net current asset" in text:
            metadata["cash"] = numbers[-1]
        elif "net current assets" in text and len(numbers) >= 2:
            # Nippon reports % of NAV as a decimal fraction in the final column.
            ratio = numbers[-1]
            metadata["cash"] = ratio * 100 if abs(ratio) <= 1 else ratio

    metadata["source_sheet"] = sheet_name
    return metadata


def extract_report_metadata(
    month: str, scheme_codes: list[str], *, research_missing: bool = False
) -> dict[str, Any]:
    """Read AMC workbooks and optionally search official factsheets for missing fields."""
    manifest_path = ROOT / "data" / "raw" / "manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    results: dict[str, Any] = {}
    for code in scheme_codes:
        snapshot = load_snapshot(code, month)
        provenance = snapshot.holdings[0].provenance
        source_file = provenance.source_file
        source_amc = provenance.amc
        source = next(
            (
                item
                for item in manifest["files"]
                if Path(item["file"]).name == source_file
                and item.get("amc") == source_amc
                and item.get("disclosure_month") == month
                and item.get("engine_eligible")
            ),
            None,
        )
        values: dict[str, Any] = {
            "aum_cr": None,
            "expense": None,
            "cat_expense": None,
            "turnover": None,
            "cash": None,
            "manager": None,
            "category": "Large Cap",
            "source": source.get("source_url") if source else source_file,
            "page_table": provenance.source_span or source_file,
            "source_file": source_file,
            "source_sheet": None,
            "disclosure_month": month,
            "field_sources": {},
        }
        if source:
            source_path = ROOT / source["file"]
            try:
                extracted = _auto_values(source_path, snapshot.scheme_name)
                values.update({key: extracted[key] for key in ("aum_cr", "expense", "cat_expense", "turnover", "cash")})
                values["source_sheet"] = extracted["source_sheet"]
            except (OSError, ValueError, ImportError, KeyError):
                # Preserve known source lineage and expose unextractable fields as unavailable.
                pass
        research_status = "not_requested"
        if research_missing:
            if not openai_lookup_enabled():
                research_status = "not_configured"
            else:
                research_status = "searched"
                suggestions = lookup_official_factsheet(
                    snapshot.scheme_name, source_amc, month
                )
                for key, citation in suggestions.items():
                    if values.get(key) is None and citation.get("value") is not None:
                        values[key] = citation["value"]
                        values["field_sources"][key] = {
                            "url": citation["source_url"],
                            "title": citation["source_title"],
                            "evidence": citation["evidence"],
                        }
        values["research_status"] = research_status
        if SCHEMES.get(code, {}).get("role") == "benchmark_control":
            values["category"] = "Large Cap Index"
        values["unavailable"] = [
            key for key in ("aum_cr", "expense", "cat_expense", "turnover", "cash", "manager")
            if values[key] is None
        ]
        results[code] = values
    return results
