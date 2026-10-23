"""CSV ingestion with hash-based identity and explicit row provenance."""

from __future__ import annotations

import csv
import hashlib
import re
from datetime import UTC, date, datetime
from decimal import Decimal
from pathlib import Path
from typing import TypedDict

import pandas as pd

from .models import Holding, PortfolioSnapshot, Provenance


class _TataHolding(TypedDict):
    name: str
    weight: Decimal
    sector: str | None
    rows: list[int]


REQUIRED = {
    "scheme_id",
    "scheme_name",
    "disclosure_month",
    "isin",
    "security_name",
    "weight_pct",
    "asset_type",
    "amc",
}


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def load_csv(path: str | Path) -> PortfolioSnapshot:
    """Load a normalized CSV; never merges or silently discards duplicate identifiers."""
    file = Path(path)
    digest = sha256_file(file)
    parsed_at = datetime.now(UTC)
    with file.open("r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames or not REQUIRED.issubset(reader.fieldnames):
            raise ValueError(f"CSV requires columns: {', '.join(sorted(REQUIRED))}")
        items = list(reader)
    if not items:
        raise ValueError("CSV contains no records")
    month = date.fromisoformat(items[0]["disclosure_month"])
    sid, sname = items[0]["scheme_id"], items[0]["scheme_name"]
    holdings = []
    for row_num, row in enumerate(items, start=2):
        if (
            date.fromisoformat(row["disclosure_month"]) != month
            or row["scheme_id"] != sid
            or row["scheme_name"] != sname
        ):
            raise ValueError(f"row {row_num}: one file must contain exactly one scheme-month")
        holdings.append(
            Holding(
                isin=row["isin"],
                security_name=row["security_name"],
                weight_pct=Decimal(row["weight_pct"]),
                asset_type=row["asset_type"].strip().lower(),
                provenance=Provenance(
                    source_file=file.name,
                    sha256=digest,
                    source_row=row_num,
                    amc=row["amc"],
                    disclosure_month=month,
                    parsed_at=parsed_at,
                    source_span=f"CSV row {row_num}",
                ),
            )
        )
    return PortfolioSnapshot(
        scheme_id=sid, scheme_name=sname, disclosure_month=month, holdings=holdings
    )


def load_hdfc_xlsx(
    path: str | Path, *, scheme_id: str, scheme_name: str, amc: str = "HDFC Mutual Fund"
) -> PortfolioSnapshot:
    """Extract a HDFC monthly workbook using its labelled ISIN and % to NAV columns.

    Row numbers in provenance are the original one-based worksheet rows. The parser
    only accepts records inside an explicitly labelled equity section; other asset
    classes are omitted until an adapter can classify their source sections safely.
    """
    file = Path(path)
    digest = sha256_file(file)
    book = pd.ExcelFile(file)
    if not book.sheet_names:
        raise ValueError("workbook contains no worksheets")
    sheet = pd.read_excel(file, sheet_name=book.sheet_names[0], header=None, dtype=object)
    date_text = next(
        (
            str(value)
            for value in sheet.iloc[:3].values.ravel()
            if isinstance(value, str) and "Portfolio as on" in value
        ),
        None,
    )
    if date_text is None:
        raise ValueError("workbook is missing its 'Portfolio as on' date")
    raw_date = date_text.partition("Portfolio as on")[2].strip()
    month = pd.to_datetime(raw_date, format="%d-%b-%Y", errors="raise").date().replace(day=1)

    header_row = next(
        (
            idx
            for idx in range(min(20, len(sheet)))
            if any(str(v).strip().upper() == "ISIN" for v in sheet.iloc[idx].values if pd.notna(v))
        ),
        None,
    )
    if header_row is None:
        raise ValueError("workbook is missing an ISIN header row")
    headers = [str(v).strip().lower() if pd.notna(v) else "" for v in sheet.iloc[header_row]]
    isin_col = next(i for i, v in enumerate(headers) if v == "isin")
    name_col = next((i for i, v in enumerate(headers) if "name of the instrument" in v), None)
    weight_col = next((i for i, v in enumerate(headers) if "% to nav" in v), None)
    if name_col is None or weight_col is None:
        raise ValueError("workbook is missing instrument-name or % to NAV columns")

    holdings: list[Holding] = []
    in_equity_section = False
    isin_pattern = re.compile(r"^[A-Z]{2}[A-Z0-9]{10}$")
    for idx in range(header_row + 1, len(sheet)):
        row = sheet.iloc[idx]
        row_text = " ".join(str(v).upper() for v in row if pd.notna(v))
        label = str(row.iloc[isin_col]).strip().upper() if pd.notna(row.iloc[isin_col]) else ""
        if "EQUITY & EQUITY RELATED" in row_text or label == "EQUITY":
            in_equity_section = True
            continue
        if any(term in row_text for term in ("DEBT INSTRUMENT", "MONEY MARKET", "DERIVATIVE")):
            in_equity_section = False
        isin = label
        name = row.iloc[name_col]
        weight = row.iloc[weight_col]
        if not in_equity_section or not isin_pattern.fullmatch(isin):
            continue
        if pd.isna(name) or pd.isna(weight):
            continue
        holdings.append(
            Holding(
                isin=isin,
                security_name=str(name).replace("�", "").strip(),
                weight_pct=Decimal(str(weight)),
                asset_type="equity",
                confidence=Decimal("0.99"),
                sector=(
                    str(row.iloc[headers.index("industry+ /rating")]).strip()
                    if "industry+ /rating" in headers
                    and pd.notna(row.iloc[headers.index("industry+ /rating")])
                    else None
                ),
                provenance=Provenance(
                    source_file=file.name,
                    sha256=digest,
                    source_row=idx + 1,
                    amc=amc,
                    disclosure_month=month,
                    parsed_at=datetime.now(UTC),
                    source_span=f"{book.sheet_names[0]}!B{idx + 1}:H{idx + 1}",
                ),
            )
        )
    if not holdings:
        raise ValueError("no equity rows matched the workbook's labelled section")
    return PortfolioSnapshot(
        scheme_id=scheme_id,
        scheme_name=scheme_name,
        disclosure_month=month,
        holdings=holdings,
    )


def load_nippon_xls(
    path: str | Path, *, scheme_id: str, scheme_name: str, amc: str = "Nippon India Mutual Fund"
) -> PortfolioSnapshot:
    """Parse an official Nippon India workbook for the configured scheme name."""
    file = Path(path)
    digest = sha256_file(file)
    book = pd.ExcelFile(file)
    index_sheet = next(
        (str(name) for name in book.sheet_names if str(name).casefold() == "index"), None
    )
    if index_sheet is None:
        raise ValueError("Nippon workbook is missing INDEX sheet")
    index = pd.read_excel(file, sheet_name=index_sheet, header=None, dtype=object)
    normalized_name = " ".join(scheme_name.upper().split())
    sheet_name = None
    for _, row in index.iterrows():
        cells = [str(v) for v in row if pd.notna(v)]
        if len(cells) >= 2 and " ".join(cells[1].upper().split()).startswith(normalized_name):
            candidate = cells[0].strip()
            if candidate in book.sheet_names:
                if sheet_name is not None:
                    raise ValueError(f"scheme name is ambiguous in Nippon workbook: {scheme_name}")
                sheet_name = candidate
    if sheet_name is None:
        raise ValueError(f"scheme not found in Nippon workbook index: {scheme_name}")
    sheet = pd.read_excel(file, sheet_name=sheet_name, header=None, dtype=object)
    date_text = next(
        (
            str(value)
            for value in sheet.iloc[:4].values.ravel()
            if isinstance(value, str) and "Monthly Portfolio Statement as on" in value
        ),
        None,
    )
    if date_text is None:
        raise ValueError("Nippon worksheet is missing its portfolio statement date")
    raw_date = date_text.partition("as on")[2].strip()
    month = pd.to_datetime(raw_date, format="%B %d,%Y", errors="raise").date().replace(day=1)
    header_row = next(
        (
            idx
            for idx in range(min(20, len(sheet)))
            if any(str(v).strip().upper() == "ISIN" for v in sheet.iloc[idx].values if pd.notna(v))
        ),
        None,
    )
    if header_row is None:
        raise ValueError("Nippon worksheet is missing an ISIN header row")
    headers = [str(v).strip().lower() if pd.notna(v) else "" for v in sheet.iloc[header_row]]
    isin_col = next(i for i, v in enumerate(headers) if v == "isin")
    name_col = next((i for i, v in enumerate(headers) if "name of the instrument" in v), None)
    weight_col = next((i for i, v in enumerate(headers) if "% to nav" in v), None)
    sector_col = next((i for i, v in enumerate(headers) if v == "industry / rating"), None)
    if name_col is None or weight_col is None:
        raise ValueError("Nippon worksheet is missing name or % to NAV columns")

    holdings: list[Holding] = []
    in_equity_section = False
    isin_pattern = re.compile(r"^[A-Z]{2}[A-Z0-9]{10}$")
    for idx in range(header_row + 1, len(sheet)):
        row = sheet.iloc[idx]
        row_text = " ".join(str(v).upper() for v in row if pd.notna(v))
        label = str(row.iloc[isin_col]).strip().upper() if pd.notna(row.iloc[isin_col]) else ""
        if "EQUITY & EQUITY RELATED" in row_text:
            in_equity_section = True
            continue
        if any(term in row_text for term in ("DEBT INSTRUMENT", "MONEY MARKET", "DERIVATIVE")):
            in_equity_section = False
        name, weight = row.iloc[name_col], row.iloc[weight_col]
        if not in_equity_section or not isin_pattern.fullmatch(label):
            continue
        if pd.isna(name) or pd.isna(weight):
            continue
        holdings.append(
            Holding(
                isin=label,
                security_name=str(name).strip(),
                weight_pct=Decimal(str(weight)) * Decimal(100),
                asset_type="equity",
                confidence=Decimal("0.99"),
                sector=(
                    str(row.iloc[sector_col]).strip()
                    if sector_col is not None and pd.notna(row.iloc[sector_col])
                    else None
                ),
                provenance=Provenance(
                    source_file=file.name,
                    sha256=digest,
                    source_row=idx + 1,
                    amc=amc,
                    disclosure_month=month,
                    parsed_at=datetime.now(UTC),
                    source_span=f"{sheet_name}!A{idx + 1}:I{idx + 1}",
                ),
            )
        )
    if not holdings:
        raise ValueError("no Nippon equity rows matched the labelled section")
    return PortfolioSnapshot(
        scheme_id=scheme_id,
        scheme_name=scheme_name,
        disclosure_month=month,
        holdings=holdings,
    )


def load_tata_xlsx(
    path: str | Path, *, scheme_id: str, scheme_name: str, amc: str = "Tata Mutual Fund"
) -> PortfolioSnapshot:
    """Extract a configured scheme from Tata's official all-schemes workbook."""
    file = Path(path)
    digest = sha256_file(file)
    book = pd.ExcelFile(file)
    index_name = next(
        (str(name) for name in book.sheet_names if str(name).casefold() == "index"), None
    )
    if index_name is None:
        raise ValueError("Tata workbook is missing Index worksheet")
    index = pd.read_excel(file, sheet_name=index_name, header=None, dtype=object)
    expected = " ".join(scheme_name.upper().split())
    sheet_name = None
    for _, row in index.iterrows():
        cells = [str(v).strip() for v in row if pd.notna(v)]
        for cell_idx, cell in enumerate(cells):
            if " ".join(cell.upper().split()) != expected:
                continue
            candidate = cells[cell_idx - 1] if cell_idx >= 1 else ""
            if candidate in book.sheet_names:
                if sheet_name is not None:
                    raise ValueError(f"scheme name is ambiguous in Tata index: {scheme_name}")
                sheet_name = candidate
    if sheet_name is None:
        raise ValueError(f"scheme not found in Tata workbook index: {scheme_name}")
    sheet = pd.read_excel(file, sheet_name=sheet_name, header=None, dtype=object)
    date_text = next(
        (
            str(value)
            for value in sheet.iloc[:20].values.ravel()
            if isinstance(value, str) and "as on" in value.strip().lower()
        ),
        None,
    )
    if date_text is None:
        raise ValueError("Tata worksheet is missing its as-on date")
    raw_date = date_text.lower().partition("as on")[2].strip()
    month = (
        pd.to_datetime(raw_date, format="mixed", dayfirst=True, errors="raise")
        .date()
        .replace(day=1)
    )
    header_row = next(
        (
            idx
            for idx in range(min(30, len(sheet)))
            if any("ISIN CODE" in str(v).strip().upper() for v in sheet.iloc[idx] if pd.notna(v))
        ),
        None,
    )
    if header_row is None:
        raise ValueError("Tata worksheet is missing an ISIN CODE header")
    headers = [str(v).strip().lower() if pd.notna(v) else "" for v in sheet.iloc[header_row]]
    isin_col = next(i for i, v in enumerate(headers) if "isin code" in v)
    name_col = next((i for i, v in enumerate(headers) if "name of the instrument" in v), None)
    weight_col = next((i for i, v in enumerate(headers) if "% to nav" in v), None)
    sector_col = next((i for i, v in enumerate(headers) if v == "industry"), None)
    if name_col is None or weight_col is None:
        raise ValueError("Tata worksheet is missing instrument-name or % to NAV columns")
    grouped: dict[str, _TataHolding] = {}
    in_equity_section = False
    isin_pattern = re.compile(r"^[A-Z]{2}[A-Z0-9]{10}$")
    for idx in range(header_row + 1, len(sheet)):
        row = sheet.iloc[idx]
        row_text = " ".join(str(v).upper() for v in row if pd.notna(v))
        label = str(row.iloc[isin_col]).strip().upper() if pd.notna(row.iloc[isin_col]) else ""
        if "EQUITY & EQUITY RELATED" in row_text or label == "EQUITY":
            in_equity_section = True
            continue
        if any(term in row_text for term in ("DEBT INSTRUMENT", "MONEY MARKET", "DERIVATIVE")):
            in_equity_section = False
        name, weight = row.iloc[name_col], row.iloc[weight_col]
        if (
            not in_equity_section
            or not isin_pattern.fullmatch(label)
            or pd.isna(name)
            or pd.isna(weight)
        ):
            continue
        item = grouped.get(label)
        if item is None:
            item = {
                "name": str(name).replace("�", "").strip(),
                "weight": Decimal(0),
                "sector": (
                    str(row.iloc[sector_col]).strip()
                    if sector_col is not None and pd.notna(row.iloc[sector_col])
                    else None
                ),
                "rows": [],
            }
            grouped[label] = item
        if item["name"].rstrip("^*#") != str(name).replace("�", "").strip().rstrip("^*#"):
            raise ValueError(f"duplicate ISIN {label} has conflicting security names")
        item["weight"] = item["weight"] + Decimal(str(weight))
        item["rows"].append(idx + 1)
    holdings = []
    for isin, item in grouped.items():
        source_rows = item["rows"]
        holdings.append(
            Holding(
                isin=isin,
                security_name=str(item["name"]).rstrip("^*#"),
                weight_pct=item["weight"],
                asset_type="equity",
                confidence=Decimal("0.95" if len(source_rows) > 1 else "0.99"),
                sector=item["sector"],
                provenance=Provenance(
                    source_file=file.name,
                    sha256=digest,
                    source_row=source_rows[0],
                    amc=amc,
                    disclosure_month=month,
                    parsed_at=datetime.now(UTC),
                    source_span=f"{sheet_name}!rows={','.join(map(str, source_rows))}",
                ),
            )
        )
    if not holdings:
        raise ValueError("no Tata equity rows matched the labelled section")
    return PortfolioSnapshot(
        scheme_id=scheme_id, scheme_name=scheme_name, disclosure_month=month, holdings=holdings
    )


def load_sbi_xlsx(
    path: str | Path, *, scheme_id: str, scheme_name: str, amc: str = "SBI Mutual Fund"
) -> PortfolioSnapshot:
    """Parse SBI's official all-schemes workbook via its Index sheet and % to AUM."""
    file = Path(path)
    digest = sha256_file(file)
    book = pd.ExcelFile(file)
    try:
        index_name = next((str(n) for n in book.sheet_names if str(n).casefold() == "index"), None)
        if index_name is None:
            raise ValueError("SBI workbook is missing Index worksheet")
        index = pd.read_excel(file, sheet_name=index_name, header=None, dtype=object)
        expected = " ".join(scheme_name.upper().split())
        sheet_name = None
        for _, row in index.iterrows():
            cells = [str(v).strip() for v in row if pd.notna(v)]
            for cell_idx, cell in enumerate(cells):
                if " ".join(cell.upper().split()) == expected:
                    candidate = cells[cell_idx - 1] if cell_idx >= 1 else ""
                    if candidate in book.sheet_names:
                        if sheet_name is not None:
                            raise ValueError(
                                f"scheme name is ambiguous in SBI index: {scheme_name}"
                            )
                        sheet_name = candidate
        if sheet_name is None:
            raise ValueError(f"scheme not found in SBI workbook index: {scheme_name}")
        sheet = pd.read_excel(file, sheet_name=sheet_name, header=None, dtype=object)
        report_date = None
        for _, row in sheet.iloc[:12].iterrows():
            for col_idx, value in enumerate(row):
                if isinstance(value, str) and "PORTFOLIO STATEMENT AS ON" in value.upper():
                    report_date = row.iloc[col_idx + 1] if col_idx + 1 < len(row) else None
                    break
            if report_date is not None:
                break
        if report_date is None or pd.isna(report_date):
            raise ValueError("SBI worksheet is missing its portfolio statement date")
        month = pd.to_datetime(report_date, errors="raise").date().replace(day=1)
        header_row = next(
            (
                idx
                for idx in range(min(20, len(sheet)))
                if any(str(v).strip().upper() == "ISIN" for v in sheet.iloc[idx] if pd.notna(v))
            ),
            None,
        )
        if header_row is None:
            raise ValueError("SBI worksheet is missing an ISIN header")
        headers = [str(v).strip().lower() if pd.notna(v) else "" for v in sheet.iloc[header_row]]
        isin_col = next(i for i, v in enumerate(headers) if v == "isin")
        name_col = next((i for i, v in enumerate(headers) if "name of the instrument" in v), None)
        weight_col = next(
            (i for i, v in enumerate(headers) if "% to aum" in v or "% to nav" in v), None
        )
        market_value_col = next(
            (i for i, v in enumerate(headers) if "market value" in v and "lakhs" in v), None
        )
        sector_col = next((i for i, v in enumerate(headers) if "industry" in v), None)
        if name_col is None or weight_col is None:
            raise ValueError("SBI worksheet is missing security-name or % to AUM columns")
        grouped: dict[str, _TataHolding] = {}
        total_aum = None
        for _, total_row in sheet.iterrows():
            total_text = " ".join(str(v).upper() for v in total_row if pd.notna(v))
            if "GRAND TOTAL (AUM)" in total_text:
                numeric_values = [
                    v for v in total_row if isinstance(v, (int, float)) and pd.notna(v)
                ]
                if numeric_values:
                    total_aum = Decimal(str(numeric_values[0]))
                break
        derived_weight_rows: set[int] = set()
        in_equity_section = False
        isin_pattern = re.compile(r"^[A-Z]{2}[A-Z0-9]{10}$")
        for idx in range(header_row + 1, len(sheet)):
            row = sheet.iloc[idx]
            row_text = " ".join(str(v).upper() for v in row if pd.notna(v))
            label = str(row.iloc[isin_col]).strip().upper() if pd.notna(row.iloc[isin_col]) else ""
            if "EQUITY & EQUITY RELATED" in row_text:
                in_equity_section = True
                continue
            if any(term in row_text for term in ("DEBT INSTRUMENT", "MONEY MARKET", "DERIVATIVE")):
                in_equity_section = False
            name, weight = row.iloc[name_col], row.iloc[weight_col]
            if (
                not in_equity_section
                or not isin_pattern.fullmatch(label)
                or pd.isna(name)
                or pd.isna(weight)
            ):
                continue
            weight_text = str(weight).strip()
            if weight_text == "#":
                if market_value_col is None or total_aum is None or total_aum <= 0:
                    raise ValueError(
                        "SBI # weight cannot be derived without market value and total AUM"
                    )
                market_value = row.iloc[market_value_col]
                if pd.isna(market_value):
                    raise ValueError(f"SBI row {idx + 1} has # weight and no market value")
                weight_value = Decimal(str(market_value)) / total_aum * Decimal(100)
                derived_weight_rows.add(idx + 1)
            else:
                try:
                    weight_value = Decimal(weight_text)
                except Exception as exc:
                    raise ValueError(
                        f"SBI row {idx + 1} has unsupported weight {weight_text!r}"
                    ) from exc
            item = grouped.get(label)
            if item is None:
                item = {
                    "name": str(name).strip(),
                    "weight": Decimal(0),
                    "sector": (
                        str(row.iloc[sector_col]).strip()
                        if sector_col is not None and pd.notna(row.iloc[sector_col])
                        else None
                    ),
                    "rows": [],
                }
                grouped[label] = item
            elif item["name"].rstrip("^*#") != str(name).strip().rstrip("^*#"):
                raise ValueError(f"duplicate ISIN {label} has conflicting security names")
            item["weight"] += weight_value
            item["rows"].append(idx + 1)
        holdings = [
            Holding(
                isin=isin,
                security_name=item["name"].rstrip("^*#"),
                weight_pct=item["weight"],
                asset_type="equity",
                confidence=Decimal("0.95" if len(item["rows"]) > 1 else "0.99"),
                sector=item["sector"],
                provenance=Provenance(
                    source_file=file.name,
                    sha256=digest,
                    source_row=item["rows"][0],
                    amc=amc,
                    disclosure_month=month,
                    parsed_at=datetime.now(UTC),
                    source_span=(
                        f"{sheet_name}!rows={','.join(map(str, item['rows']))}"
                        + (
                            "; weight derived as market value / GRAND TOTAL (AUM) * 100"
                            if derived_weight_rows.intersection(item["rows"])
                            else ""
                        )
                    ),
                ),
            )
            for isin, item in grouped.items()
        ]
        if not holdings:
            raise ValueError("no SBI equity rows matched the labelled section")
        return PortfolioSnapshot(
            scheme_id=scheme_id, scheme_name=scheme_name, disclosure_month=month, holdings=holdings
        )
    finally:
        book.close()
