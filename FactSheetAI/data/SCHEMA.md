# Normalized monthly portfolio CSV schema

Columns (exact names required): `scheme_id`, `scheme_name`, `disclosure_month` (ISO date), `isin` (12 alphanumeric characters), `security_name`, `weight_pct` (decimal percent of NAV), `asset_type` (`equity`, `cash`, `debt`, `derivative`, etc.), `amc`.

One source file must represent one scheme and one month. Duplicate equity ISIN rows are rejected for human reconciliation. These files are normalized inputs, not original disclosures; preserve originals separately with their own hashes. Keep the source-to-normalized mapping and review uncertainty before interpreting results.
