# Artifact traceability

| Artifact | Module path | Automated evidence | Status |
|---|---|---|---|
| A1 Structured extraction | `factsheetai/ingestion.py`; `factsheetai/models.py` | `tests/test_ingestion.py`: normalized CSV contracts; 12 months each of HDFC/Nippon/Tata/SBI sources; hashes, source spans, duplicate and malformed-input behavior | Working source adapters; no OCR, gold set, measured field precision/recall, or review queue |
| A2 Retrieval pipeline | — | No retrieval/evaluation tests | Not implemented; no embeddings, BM25, reranking, grounded Q&A, or 30-question eval |
| A7 Governance note | `GOVERNANCE.md` | Documented model inventory, lineage, risk register, controls, oversight and limitations | Documented; monitoring and operational controls remain partial |
| A8 Fund administration | `factsheetai/models.py`, `factsheetai/catalog.py`, `factsheetai/analytics.py` | `tests/test_analytics.py`, `tests/test_dynamic.py`, `tests/test_ingestion.py`; ISIN identities, effective AMFI memberships, portfolio formulas | ISIN-based holdings and analytics implemented; full NAV/AUM reconciliation and non-equity ledger incomplete |
| A9 Dashboard/commentary | `factsheetai/api.py`, `factsheetai/static/index.html` | `tests/test_api.py`; source-backed dynamic pair analysis and provenance payload | Static dashboard and REST implemented; no automated natural-language commentary, trend/matrix view, or browser-console evidence yet |

## Evidence pointers

- Original bytes and official source URLs: `data/raw/` and `data/raw/source_index.json`.
- Hash and source-size evidence: `data/raw/manifest.json`.
- Exact AMFI Direct Growth mapping: `data/scheme-resolution.json`.
- Missing source evidence: `data/missing-source-files.json`.
- Analysis formulas and sample limits: `METHODOLOGY.md`.
- Quality commands and current milestone gates: `PROGRESS.md`.
