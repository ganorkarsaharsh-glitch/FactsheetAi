# FactSheetAI project instructions

## Purpose
FactSheetAI is a dynamic, deterministic mutual fund portfolio overlap and mandate-drift analysis project. Never invent scheme identifiers, holdings, weights, NAV/AUM, dates, or source evidence. Missing source material must remain explicitly unavailable.

## Conventions
- Python 3.11+; typed, documented functions; Decimal arithmetic where financial values are involved.
- ISIN is the primary security key. Preserve source-file SHA-256 and row-level provenance.
- Calculations are deterministic and independent of any language model.
- Schemes, thresholds, and mandates belong in versioned config, not in calculation code.
- Raw source data is immutable. Do not add fabricated or placeholder observations.

## Commands
- Install: `python -m pip install -e ".[dev]"`
- Lint: `ruff check .`
- Typecheck: `mypy factsheetai`
- Test: `pytest --cov=factsheetai --cov-report=term-missing`
- Build: `python -m build`
- Run: `python -m factsheetai --help`

## Anti-hallucination contract
Every displayed number must derive from loaded records with provenance. Missing values are reported as unavailable. Never infer or fabricate source facts. Reject unsupported questions. If a source is unavailable, document the specific limitation and continue independent work.
