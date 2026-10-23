"""Readers for official AMFI NAV and half-yearly capitalisation files."""

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from .ingestion import sha256_file


@dataclass(frozen=True)
class Top100List:
    effective_date: str
    isins: frozenset[str]
    source_file: str
    sha256: str


def load_top100(path: str | Path, effective_date: str) -> Top100List:
    """Read the 100 highest ranked companies from an official AMFI workbook."""
    file = Path(path)
    raw = pd.read_excel(file, header=None, dtype=object)
    header_idx = next(
        (
            idx
            for idx in range(min(8, len(raw)))
            if any(str(v).strip().upper() == "ISIN" for v in raw.iloc[idx] if pd.notna(v))
        ),
        None,
    )
    if header_idx is None:
        raise ValueError("AMFI workbook is missing an ISIN header")
    headers = [str(v).strip().lower() if pd.notna(v) else "" for v in raw.iloc[header_idx]]
    isin_col = next(i for i, v in enumerate(headers) if v == "isin")
    category_col = next((i for i, v in enumerate(headers) if "categorization" in v), None)
    cap_col = next((i for i, v in enumerate(headers) if "average of all exchanges" in v), None)
    if category_col is None or cap_col is None:
        raise ValueError("AMFI workbook lacks market-cap ranking columns")
    rows: list[tuple[float, str, str]] = []
    for _, row in raw.iloc[header_idx + 1 :].iterrows():
        isin_value, category, cap = row.iloc[isin_col], row.iloc[category_col], row.iloc[cap_col]
        if pd.isna(isin_value) or pd.isna(category) or pd.isna(cap):
            continue
        isin = str(isin_value).strip().upper()
        if len(isin) == 12 and str(category).strip().casefold() == "large cap":
            rows.append((float(cap), isin, str(category).strip()))
    top = sorted(rows, key=lambda item: -item[0])[:100]
    if len(top) != 100:
        raise ValueError(f"AMFI workbook contains only {len(top)} valid large-cap records")
    return Top100List(
        effective_date=effective_date,
        isins=frozenset(item[1] for item in top),
        source_file=file.name,
        sha256=sha256_file(file),
    )
