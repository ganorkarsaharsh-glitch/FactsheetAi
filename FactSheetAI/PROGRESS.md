# FactSheetAI progress

| Milestone | Status | Gate result | Open issues | Next step |
|---|---|---|---|---|
| M0 Setup | Partial | Bundled Python 3.12 runtime passes compile, lint, typecheck, and unit tests | Standard pytest and Docker/CI are not configured; package installation from PyPI is blocked by DNS | Add reproducible CI and one-command deployment |
| M1 Data acquisition | Partial | 188,908,286 bytes of archived AMC/AMFI sources cover five schemes and 12 months, with AMFI effective-dated lists | ICICI holding ISINs and several requested AMC sources remain unavailable | Acquire missing official AMC source disclosures |
| M2 Extraction/admin | Partial | Normalized CSV parser and non-persisting CSV review preview with SHA-256 and row provenance | No PDF/XLS disclosure extraction, editable review/import workflow, or labelled gold set | Add adapters and review workflow against approved originals |
| M3 Engine | Partial | N-fund overlap, three-path validation, weighted look-through, concentration, holding/sector drift, cash-aware active share, configurable red flags and health score | No source-backed cash/expense/AUM/manager fields or official benchmark constituent data | Acquire verified metadata and benchmark inputs |
| M4 Validation | Partial | Standard-library analytics, API, CSV preview, and formula-template tests; hand-calculated overlap/active share/HHI/drift examples | No third-party overlap baseline or source-backed golden labels; coverage target not measured | Add independent source reconciliation and coverage reporting |
| M5 Retrieval/LLM | Not started | Deterministic metric-grounded plain-English summary and overlap-only simulation are available | No source-indexed document corpus, LLM credentials, citation evaluation set, or chat | Implement only after source corpus and citation evaluation are available |
| M6 API | Partial | Catalog, pair and 2–4 fund analysis, threshold edit, upload preview, and validated Excel export routes | No authentication, rate limits, database, or live verified LibreOffice runtime | Verify Excel recalc in a host with LibreOffice; add deployment security |
| M7 Scheduler/admin | Not started | — | No AMC endpoint inventory or network | Add authorized per-AMC adapters and job controls |
| M8 Dashboard | Partial | Responsive static UI with pair mode, multi-fund heatmap, allocations, health, drift, threshold editor, local save/share, and CSV preview | Existing FastAPI/static stack retained; no Framer Motion/21st.dev, PDF upload, or browser visual QA | Run UI in browser and iterate against viva device |
| M9 Advanced features | Partial | Excel template integrated with formula scan, swap simulation, and threshold editing | Excel delivery requires user-reviewed metadata, a disclosed index-fund proxy, and LibreOffice; PDF summary not available | Supply verified metadata and validate with LibreOffice |
| M10 Demo | Partial | Four source-available large-cap disclosures can be compared; three overlap recalculations and a viva cheat sheet are included | Requested ICICI holding-level source remains unavailable; current four-fund demo includes Tata peer, not the missing ICICI scheme | Keep the limitation explicit until source records arrive |
| M11 Governance/docs | Partial | Initial methodology and governance risk controls documented | No sourced research findings, limitations register completeness, viva Q&A, or evidence corpus | Complete with verified snapshots and actual evaluated system |

## Commands and results
- `python -m compileall -q factsheetai tests` — passed using bundled Python 3.12.14.
- `python -m unittest discover -s tests -q` — 32 tests passed using bundled Python 3.12.14.
- `python -m pytest --cov=factsheetai --cov-report=term-missing` — unavailable: bundled runtime has no pytest.
- `python -m pip install -e ".[dev]"` — blocked: DNS resolution for `pypi.org` failed; sandbox network access is unavailable. No dependencies were downloaded.
- `python -m factsheetai --help` — passed using bundled Python 3.12.14.
- `ruff check .` — passed.
- `mypy factsheetai` — passed (9 source files).
- Inline dashboard JavaScript syntax check with bundled Node — passed. No browser screenshot or Lighthouse run was performed.
