# Model governance and AI governance note

## Purpose, users and intended use

FactSheetAI is a student research and viva prototype for comparing monthly disclosed equity holdings by ISIN. Its users are the student, evaluator and research audience. It is educational analysis and is not intended for investment recommendations, suitability decisions, execution, or regulatory reporting.

## Model inventory

| Component | Role | Decision authority |
|---|---|---|
| AMC workbook adapters and normalized CSV loader | Parse holding identifiers, names, weights and source spans | Schema-validated deterministic extraction; humans must review before accounting use |
| AMFI top-100 reader | Rank large-cap rows and create effective-dated membership | Deterministic; source workbook and effective date retained |
| Analytics engine | Overlap, Jaccard, concentration, drift, turnover, active share and red flags | Deterministic Decimal/Python code, with pandas overlap cross-check |
| Language model | None in this release | No generated analysis or prose claims |

The original request's retrieval/LLM layer is not implemented. This avoids unsupported generated claims. If added later, it must answer only from retrieved source spans plus computed JSON, attach citations to every sentence, reject unsupported questions and post-check every numeric token.

## Data lineage and audit

`data/raw/source_index.json` records official source URLs, disclosure month and source role. `data/raw/manifest.json` records byte size and SHA-256. Parsed snapshots retain scheme code, source filename, hash, AMC, month, parse timestamp, source row/page span and field confidence. A monthly result is reproducible from its immutable source file, parser version and effective AMFI list. Raw files are not normalized in place.

## Risk register and controls

| Risk | Control in this release | Remaining risk |
|---|---|---|
| Fabricated or mistaken security identity | ISIN-first joins; no fuzzy match; missing data remains unavailable | A source may itself contain erroneous ISINs |
| Parser layout changes or weight conversion error | Per-AMC adapters; strict section/header detection; malformed values fail; source hash/span | No labeled 5-scheme gold set or parser recall/precision estimate |
| Duplicate/conflicting holdings | Sum same-ISIN rows only with matching normalized names; reject conflicts | Other AMC-specific duplicate semantics need review |
| SBI marker substitution | Derive marked percentage only from same-row market value and same-workbook total AUM; provenance records formula | Needs a source-review signoff and separate accounting reconciliation |
| Stale or mismatched market-cap list | Explicit half-year effective-date configuration, matched to disclosure month | Future list records must be verified before their effective period |
| False mandate alert | Compare only reported equity holdings to effective AMFI top 100; show source date | Derivatives, cash, and parser omissions affect interpretation |
| Calculation defect | Decimal metrics, independent pandas/Python overlap, property-based tests | No published third-party benchmark or spreadsheet parity test |
| Missing key fund | Catalog and report list missing months; API errors rather than substitute | ICICI prevents answering the original four-fund question |
| LLM hallucination | No LLM in calculation or UI | Future LLM requires citation and number-check evaluation |
| Scraper/download breakage | Archived immutable files and source manifest | No automatic scheduler, retries, or live ingestion monitoring |

## Human oversight, monitoring and change management

This release has no operator review queue, auth, audit-event store, or live source monitor. The source manifest and checked-in config provide a basic review trail. A reviewer should verify a parser against at least five hand-labeled schemes and reconcile reported equity totals to AMC NAV/AUM before using results beyond demonstration. Changes to adapter logic, mandate thresholds or effective dates require a reviewed code/config change, a manifest rebuild and full regression run. Ingestion errors must remain visible and must not be converted to synthetic or inferred observations.

## Limitations

Only four large-cap schemes and one index control have 12 months of ISIN-level files. ICICI has factsheets but not complete ISIN holdings; other configured schemes lack source files. NAV/AUM reconciliation, security-class decomposition beyond equity, precision/recall evaluation, retrieval, user workspace, auth, scheduler and operations logging are incomplete. The raw source corpus is 188,908,286 bytes, below the requested 300 MB target. The current prototype is not accounting, compliance or investment advice software.
