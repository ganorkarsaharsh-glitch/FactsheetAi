"""Deterministic calculations over source-backed snapshot records."""

from collections.abc import Mapping
from decimal import Decimal
from typing import TypedDict

import pandas as pd

from .models import Holding, PortfolioSnapshot

D = Decimal


class OverlapResult(TypedDict):
    overlap_pct: Decimal
    jaccard: Decimal
    common_isins: list[str]


class TopHoldingsResult(TypedDict):
    top_n: list[Holding]
    top_n_concentration_pct: Decimal
    equity_hhi: Decimal
    effective_stocks: Decimal
    equity_weight_pct: Decimal


class HoldingDriftRow(TypedDict):
    isin: str
    name: str
    previous_pct: Decimal
    current_pct: Decimal
    change_pct: Decimal
    status: str


class SectorDriftRow(TypedDict):
    sector: str
    previous_pct: Decimal
    current_pct: Decimal
    change_pct: Decimal


DEFAULT_LARGE_CAP_FLOOR = D(80)
DEFAULT_TOP10_ALERT_THRESHOLD = D(50)


def equity_weights(snapshot: PortfolioSnapshot) -> dict[str, Decimal]:
    return {h.isin: h.weight_pct for h in snapshot.holdings if h.asset_type == "equity"}


def overlap(a: PortfolioSnapshot, b: PortfolioSnapshot) -> OverlapResult:
    """Shared NAV weight is sum(min(weight A, weight B)) for common ISINs."""
    aw, bw = equity_weights(a), equity_weights(b)
    common = sorted(aw.keys() & bw.keys())
    shared = sum((min(aw[i], bw[i]) for i in common), D(0))
    union = aw.keys() | bw.keys()
    return {
        "overlap_pct": shared,
        "jaccard": D(len(common)) / D(len(union)) if union else D(0),
        "common_isins": common,
    }


def overlap_matrix(snapshots: Mapping[str, PortfolioSnapshot]) -> dict[str, dict[str, Decimal]]:
    """Symmetric pairwise matrix. Diagonal is actual equity weight, capped at 100."""
    keys = list(snapshots)
    out: dict[str, dict[str, Decimal]] = {k: {} for k in keys}
    for i, a in enumerate(keys):
        for b in keys[i:]:
            val = (
                min(D(100), sum(equity_weights(snapshots[a]).values(), D(0)))
                if a == b
                else overlap(snapshots[a], snapshots[b])["overlap_pct"]
            )
            out[a][b] = val
            out[b][a] = val
    return out


def top_holdings_stats(snapshot: PortfolioSnapshot, n: int = 10) -> TopHoldingsResult:
    rows = sorted(
        (h for h in snapshot.holdings if h.asset_type == "equity"),
        key=lambda h: (-h.weight_pct, h.isin),
    )
    total = sum((h.weight_pct for h in rows), D(0))
    hhi = sum(((h.weight_pct / D(100)) ** 2 for h in rows), D(0))
    return {
        "top_n": rows[:n],
        "top_n_concentration_pct": sum((h.weight_pct for h in rows[:n]), D(0)),
        "equity_hhi": hhi,
        "effective_stocks": D(1) / hhi if hhi else D(0),
        "equity_weight_pct": total,
    }


def style_drift(
    snapshot: PortfolioSnapshot,
    top100_isins: set[str],
    large_cap_floor_pct: Decimal = DEFAULT_LARGE_CAP_FLOOR,
) -> dict[str, object]:
    eq = [h for h in snapshot.holdings if h.asset_type == "equity"]
    outside = sum((h.weight_pct for h in eq if h.isin not in top100_isins), D(0))
    inside = sum((h.weight_pct for h in eq if h.isin in top100_isins), D(0))
    return {
        "top100_weight_pct": inside,
        "outside_top100_weight_pct": outside,
        "large_cap_floor_pct": large_cap_floor_pct,
        "mandate_breach": inside < large_cap_floor_pct,
    }


def pandas_overlap(a: PortfolioSnapshot, b: PortfolioSnapshot) -> Decimal:
    left = pd.DataFrame(
        [(h.isin, h.weight_pct) for h in a.holdings if h.asset_type == "equity"],
        columns=["isin", "a"],
    )
    right = pd.DataFrame(
        [(h.isin, h.weight_pct) for h in b.holdings if h.asset_type == "equity"],
        columns=["isin", "b"],
    )
    joined = left.merge(right, on="isin", how="inner")
    return sum((min(x, y) for x, y in zip(joined["a"], joined["b"])), D(0))


def python_overlap(a: PortfolioSnapshot, b: PortfolioSnapshot) -> Decimal:
    left, right = equity_weights(a), equity_weights(b)
    return sum((min(left[k], right[k]) for k in left.keys() & right.keys()), D(0))


def weighted_portfolio(
    snapshots: Mapping[str, PortfolioSnapshot], allocations_pct: Mapping[str, Decimal]
) -> dict[str, Decimal]:
    """Return stock exposures for a user-weighted mix; allocations are normalized."""
    if not allocations_pct or set(allocations_pct) - set(snapshots):
        raise ValueError("each allocation must refer to a loaded scheme")
    if any(weight < 0 for weight in allocations_pct.values()):
        raise ValueError("allocations cannot be negative")
    total = sum(allocations_pct.values(), D(0))
    if total <= 0:
        raise ValueError("allocation total must be positive")
    combined: dict[str, Decimal] = {}
    for scheme, allocation in allocations_pct.items():
        for isin, weight in equity_weights(snapshots[scheme]).items():
            combined[isin] = combined.get(isin, D(0)) + allocation / total * weight
    return combined


def portfolio_concentration(
    snapshots: Mapping[str, PortfolioSnapshot], allocations_pct: Mapping[str, Decimal]
) -> dict[str, Decimal]:
    """Concentration of a blended investor portfolio, normalized to equity exposure."""
    combined = weighted_portfolio(snapshots, allocations_pct)
    equity_total = sum(combined.values(), D(0))
    fractions = [weight / equity_total for weight in combined.values()] if equity_total else []
    hhi = sum((fraction**2 for fraction in fractions), D(0))
    return {
        "unique_stocks": D(len(combined)),
        "equity_exposure_pct": equity_total,
        "hhi": hhi,
        "effective_stocks": D(1) / hhi if hhi else D(0),
    }


def marginal_overlap(
    existing: Mapping[str, PortfolioSnapshot],
    existing_allocations_pct: Mapping[str, Decimal],
    candidate: PortfolioSnapshot,
) -> Decimal:
    """Overlap of a candidate with the existing blended portfolio, by min weight."""
    blend = weighted_portfolio(existing, existing_allocations_pct)
    candidate_weights = equity_weights(candidate)
    return sum((min(w, candidate_weights.get(i, D(0))) for i, w in blend.items()), D(0))


def turnover(previous: PortfolioSnapshot, current: PortfolioSnapshot) -> Decimal:
    """One-way month-to-month turnover: half the absolute weight changes."""
    before, after = equity_weights(previous), equity_weights(current)
    return sum(
        (abs(before.get(i, D(0)) - after.get(i, D(0))) for i in before.keys() | after.keys()), D(0)
    ) / D(2)


def active_share(portfolio: PortfolioSnapshot, benchmark: PortfolioSnapshot) -> Decimal:
    """Half the NAV-weight difference across equities and cash, including the cash bucket."""
    def weights(snapshot: PortfolioSnapshot) -> dict[str, Decimal]:
        out = {h.isin: h.weight_pct for h in snapshot.holdings if h.asset_type == "equity"}
        cash = sum((h.weight_pct for h in snapshot.holdings if h.asset_type == "cash"), D(0))
        if cash:
            out["__CASH__"] = cash
        return out
    a, b = weights(portfolio), weights(benchmark)
    return sum((abs(a.get(i, D(0)) - b.get(i, D(0))) for i in a.keys() | b.keys()), D(0)) / D(2)


def identity_overlap(a: PortfolioSnapshot, b: PortfolioSnapshot) -> Decimal:
    """Independent identity using min(a,b) = (a+b-|a-b|)/2 for common ISINs."""
    left, right = equity_weights(a), equity_weights(b)
    return sum(((left[i] + right[i] - abs(left[i] - right[i])) / D(2)
                for i in left.keys() & right.keys()), D(0))


def overlap_rank(value: Decimal) -> str:
    return "HIGH" if value >= D(50) else "MODERATE" if value >= D(30) else "LOW"


def holding_drift(previous: PortfolioSnapshot, current: PortfolioSnapshot) -> dict[str, object]:
    """Source-keyed monthly position changes, with new/exited names made explicit."""
    before, after = equity_weights(previous), equity_weights(current)
    old_names = {h.isin: h.security_name for h in previous.holdings}
    new_names = {h.isin: h.security_name for h in current.holdings}
    rows: list[HoldingDriftRow] = [{"isin": i, "name": new_names.get(i, old_names.get(i, i)),
             "previous_pct": before.get(i, D(0)), "current_pct": after.get(i, D(0)),
             "change_pct": after.get(i, D(0)) - before.get(i, D(0)),
             "status": "NEW" if i not in before else "EXITED" if i not in after else "CHANGED"}
            for i in sorted(before.keys() | after.keys())]
    rows.sort(key=lambda r: (-abs(r["change_pct"]), r["isin"]))
    return {"rows": rows, "new_entries": [r for r in rows if r["status"] == "NEW"],
            "exits": [r for r in rows if r["status"] == "EXITED"],
            "biggest_increases": sorted((r for r in rows if r["change_pct"] > 0), key=lambda r: -r["change_pct"])[:10],
            "biggest_decreases": sorted((r for r in rows if r["change_pct"] < 0), key=lambda r: r["change_pct"])[:10],
            "portfolio_change_pct": sum((abs(r["change_pct"]) for r in rows), D(0)) / D(2)}


def sector_drift(previous: PortfolioSnapshot, current: PortfolioSnapshot) -> list[SectorDriftRow]:
    def weights(snapshot: PortfolioSnapshot) -> dict[str, Decimal]:
        values: dict[str, Decimal] = {}
        for holding in snapshot.holdings:
            if holding.asset_type == "equity" and holding.sector:
                values[holding.sector] = values.get(holding.sector, D(0)) + holding.weight_pct
        return values
    before, after = weights(previous), weights(current)
    rows: list[SectorDriftRow] = [{"sector": sector, "previous_pct": before.get(sector, D(0)),
                    "current_pct": after.get(sector, D(0)),
                    "change_pct": after.get(sector, D(0)) - before.get(sector, D(0))}
                   for sector in before.keys() | after.keys()]
    return sorted(rows, key=lambda row: (-abs(row["change_pct"]), row["sector"]))


def concentration_detail(snapshot: PortfolioSnapshot) -> dict[str, object]:
    stats = top_holdings_stats(snapshot, 10)
    equities = [h for h in snapshot.holdings if h.asset_type == "equity"]
    sectors: dict[str, Decimal] = {}
    for h in equities:
        if h.sector:
            sectors[h.sector] = sectors.get(h.sector, D(0)) + h.weight_pct
    equity_weight = sum((h.weight_pct for h in equities), D(0))
    classified_weight = sum(sectors.values(), D(0))
    return {"top5_pct": sum((h.weight_pct for h in stats["top_n"][:5]), D(0)),
            "top10_pct": stats["top_n_concentration_pct"], "hhi": stats["equity_hhi"],
            "effective_stocks": stats["effective_stocks"],
            "largest_holding": max(equities, key=lambda h: h.weight_pct).weight_pct if equities else D(0),
            "largest_sector": max(sectors, key=lambda sector: sectors[sector]) if sectors else None,
            "largest_sector_pct": max(sectors.values()) if sectors else D(0),
            "sector_coverage_pct": classified_weight / equity_weight * D(100) if equity_weight else D(0)}


def combined_portfolio_detail(snapshots: Mapping[str, PortfolioSnapshot],
                              investments: Mapping[str, Decimal]) -> dict[str, object]:
    exposures: dict[str, Decimal] = {}
    total_investment = sum(investments.values(), D(0))
    if not investments or total_investment <= 0:
        raise ValueError("investment amounts must total more than zero")
    names: dict[str, str] = {}
    for fund, amount in investments.items():
        if fund not in snapshots or amount < 0:
            raise ValueError("investments must be non-negative and refer to loaded funds")
        if amount == 0:
            continue
        for holding in snapshots[fund].holdings:
            if holding.asset_type == "equity" and holding.weight_pct > 0:
                exposures[holding.isin] = exposures.get(holding.isin, D(0)) + amount * holding.weight_pct / D(100)
                names[holding.isin] = holding.security_name
    invested_equity = sum(exposures.values(), D(0))
    fractions = [v / invested_equity for v in exposures.values()] if invested_equity else []
    hhi = sum((v*v for v in fractions), D(0))
    ordered = sorted(exposures.items(), key=lambda x: (-x[1], x[0]))
    stock_positions = sum(len(equity_weights(snapshots[k])) for k, amount in investments.items() if amount > 0)
    return {"unique_stocks": len(exposures), "stock_positions": stock_positions,
            "duplication_ratio": D(1) - D(len(exposures)) / D(max(1, stock_positions)),
            "effective_stocks": D(1) / hhi if hhi else D(0),
            "top10_share_pct": sum((v for _, v in ordered[:10]), D(0)) / invested_equity * D(100) if invested_equity else D(0),
            "exposures_rupees": {names[k]: v for k, v in ordered}, "equity_rupees": invested_equity,
            "investment_rupees": total_investment}


def validation_detail(snapshots: Mapping[str, PortfolioSnapshot], benchmark: PortfolioSnapshot | None = None) -> dict[str, object]:
    checks = []
    for key, snap in snapshots.items():
        eq = sum((h.weight_pct for h in snap.holdings if h.asset_type == "equity"), D(0))
        cash = sum((h.weight_pct for h in snap.holdings if h.asset_type == "cash"), D(0))
        checks.append({"fund": key, "equity_pct": eq, "cash_pct": cash,
                       "weights_cash_sum_pct": eq + cash,
                       "weights_cash_approximately_100": abs(eq + cash - D(100)) <= D("0.5"),
                       "duplicate_isins": len([h.isin for h in snap.holdings if h.asset_type == "equity"]) - len(set(equity_weights(snap))),
                       "benchmark_unmatched_isins": sorted(set(equity_weights(snap)) - set(equity_weights(benchmark))) if benchmark else []})
    all_names: dict[str, set[str]] = {}
    for snap in snapshots.values():
        for h in snap.holdings:
            if h.asset_type == "equity":
                all_names.setdefault(h.isin, set()).add(" ".join(h.security_name.casefold().split()))
    mismatches = sorted(i for i, names in all_names.items() if len(names) > 1)
    return {"checks": checks, "all_weight_checks_pass": all(c["weights_cash_approximately_100"] for c in checks),
            "cross_fund_name_mismatches": mismatches,
            "note": "Equity-only source adapters do not include cash; weight checks can therefore be incomplete. ISIN identity is authoritative; name mismatches are listed for review."}


def assess_rule(rule: Mapping[str, object], value: Decimal | None) -> str:
    """Apply configurable red limits and a 15% amber band in the appropriate direction."""
    if value is None:
        return "UNAVAILABLE"
    threshold = D(str(rule["threshold"]))
    direction = str(rule.get("direction", "above"))
    margin = D(str(rule.get("amber_margin_pct", 15))) / D(100)
    if rule.get("inclusive") and direction == "above" and value >= threshold:
        return "RED"
    if direction == "below":
        if value < threshold:
            return "RED"
        if value < threshold * (D(1) + margin):
            return "AMBER"
        return "OK"
    if value > threshold:
        return "RED"
    if value >= threshold * (D(1) - margin):
        return "AMBER"
    return "OK"


def health_score(flags: list[dict[str, object]], rules: Mapping[str, Mapping[str, object]]) -> dict[str, object]:
    score = D(100)
    for flag in flags:
        penalty = D(str(rules.get(str(flag["code"]), {}).get("penalty", 0)))
        score -= penalty if flag["severity"] == "RED" else penalty / D(2) if flag["severity"] == "AMBER" else D(0)
    score = max(D(0), score)
    return {"score": score, "band": "HEALTHY" if score >= 80 else "WATCH" if score >= 60 else "CONCERN"}


def red_flags(
    snapshot: PortfolioSnapshot,
    *,
    top100_isins: set[str] | None = None,
    previous: PortfolioSnapshot | None = None,
    max_top10_pct: Decimal = DEFAULT_TOP10_ALERT_THRESHOLD,
) -> list[dict[str, str]]:
    """Deterministic rules only; each alert points to source-backed records."""
    alerts: list[dict[str, str]] = []
    concentration = top_holdings_stats(snapshot)
    if concentration["top_n_concentration_pct"] > max_top10_pct:
        alerts.append(
            {"code": "TOP10_CONCENTRATION", "value": str(concentration["top_n_concentration_pct"])}
        )
    if top100_isins is not None:
        drift = style_drift(snapshot, top100_isins)
        if drift["mandate_breach"]:
            alerts.append({"code": "LARGE_CAP_FLOOR", "value": str(drift["top100_weight_pct"])})
    if previous is not None:
        value = turnover(previous, snapshot)
        if value >= D(10):
            alerts.append({"code": "MONTHLY_TURNOVER_SPIKE", "value": str(value)})
    return alerts
