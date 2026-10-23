# Three-minute viva walkthrough

This script uses two available source-backed schemes (HDFC Large Cap and Nippon India Large Cap) for July 2025. It demonstrates the engine but does **not** answer the full four-requested-fund research question because ICICI Prudential lacks complete ISIN-level files in the corpus.

- **0:00–0:25 — Research question and scope.** “I want to measure duplicated equity exposure across large-cap mutual funds. I compare the same disclosure month and join holdings by stock ISIN. The complete requested four-fund set is not yet available because ICICI's archived factsheets aggregate some positions and omit their ISINs.”
- **0:25–0:55 — Select a source-backed pair.** In the app choose HDFC Large Cap and Nippon India Large Cap, July 2025. Point to source file, AMC, month, SHA-256 and parsed rows. Each number can be traced back to an original source row.
- **0:55–1:30 — Explain overlap.** “For every ISIN held by both funds, I take the lower of the two portfolio weights and add those minima. This is shared portfolio weight, not correlation.” Show common ISINs and the largest contribution rows.
- **1:30–2:00 — Explain concentration.** Show top-ten weight and HHI. “Top ten sums the ten largest disclosed equity weights. HHI sums squared NAV fractions; its reciprocal is the number of equally weighted stocks that would produce the same concentration.”
- **2:00–2:25 — Explain mandate drift.** Show each scheme's exposure in the AMFI top-100 list effective for July 2025 (30 June 2025). Explain that the app does not use one current list for every month and tests the 80% large-cap floor.
- **2:25–2:45 — Validation and what-if.** Show the three overlap paths: pandas join, pure Python mapping, and the identity `(a+b-|a-b|)/2`. Then change an investment amount or remove a scheme and point out that look-through rupee exposure updates while pairwise overlap stays allocation-independent.
- **2:45–3:00 — Close honestly.** “These results describe reported holdings only and are educational. I cannot state the four-fund overlap until ICICI supplies a complete ISIN-bearing monthly portfolio. No data was guessed to fill this gap.”

## Formulas to explain

- Weighted overlap: `Σ min(wAᵢ,wBᵢ)` for common equity ISINs. The equivalent identity is `Σ (wAᵢ+wBᵢ-|wAᵢ-wBᵢ|)/2`.
- Active share, only when a valid benchmark is supplied: `0.5 × Σ|w_fund−w_benchmark|`, including the cash bucket. It is unavailable in the current corpus.
- HHI: `Σ(weight_i/100)²`; effective stocks: `1/HHI`.
- Monthly portfolio change: `0.5 × Σ|current_i−previous_i|`, including entries and exits as zero in the missing month.
- Combined holding exposure: `Σ(investment_fund × disclosed_weight_i/100)`; combined HHI is computed after normalizing these rupee exposures.
- Red flags apply configurable thresholds. The amber band is 15% of the threshold; health is 100 less full red penalties and half amber penalties, floored at zero.

The independent overlap recomputations and hand-calculated unit examples validate the arithmetic. They do not validate whether the source disclosure itself is complete or whether a security was classified correctly. Missing expense, cash, manager, AUM, and official benchmark records remain unavailable.

## Demo commands

```powershell
uvicorn factsheetai.api:app --host 127.0.0.1 --port 8000
```

Then open `http://127.0.0.1:8000/`. The app has a source-backed offline dataset, but this release does not include a one-click Viva Demo button. Plan to show a previously checked response if the live source selection or browser is unavailable.
