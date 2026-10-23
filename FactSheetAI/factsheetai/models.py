"""Validated source and normalized portfolio records."""

from datetime import date, datetime
from decimal import Decimal

from pydantic import BaseModel, Field, field_validator


class Provenance(BaseModel):
    source_file: str
    sha256: str = Field(pattern=r"^[a-f0-9]{64}$")
    source_row: int = Field(ge=1)
    amc: str
    disclosure_month: date
    parsed_at: datetime
    source_span: str | None = None


class Holding(BaseModel):
    isin: str
    security_name: str
    weight_pct: Decimal = Field(ge=0, le=100)
    asset_type: str = "equity"
    confidence: Decimal = Field(default=Decimal(1), ge=0, le=1)
    sector: str | None = None
    provenance: Provenance

    @field_validator("isin")
    @classmethod
    def valid_isin_shape(cls, value: str) -> str:
        value = value.strip().upper()
        if len(value) != 12 or not value.isalnum():
            raise ValueError("ISIN must be 12 alphanumeric characters")
        return value


class PortfolioSnapshot(BaseModel):
    scheme_id: str
    scheme_name: str
    disclosure_month: date
    holdings: list[Holding]

    @field_validator("holdings")
    @classmethod
    def unique_equity_isins(cls, rows: list[Holding]) -> list[Holding]:
        ids = [h.isin for h in rows if h.asset_type == "equity"]
        if len(ids) != len(set(ids)):
            raise ValueError("duplicate equity ISINs must be reconciled before loading")
        return rows
