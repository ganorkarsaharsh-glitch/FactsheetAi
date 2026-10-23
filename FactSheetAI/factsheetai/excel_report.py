#!/usr/bin/env python3
"""
FactSheetAI - Excel report generator
====================================
Takes the comparison result (JSON) and writes a multi-tab, FORMULA-DRIVEN
Excel workbook. Raw holdings are the only hard-coded data; every metric
(overlap, concentration, active share, drift, red flags, health score,
plain-English summary) is an Excel formula, so the user can edit inputs and
the whole workbook recalculates.

Usage
-----
  python factsheetai_excel_report.py --dump-sample sample_input.json
  python factsheetai_excel_report.py --input sample_input.json --output report.xlsx
  python factsheetai_excel_report.py            # builds report from built-in SAMPLE data

JSON schema (weights are FRACTIONS of NAV, e.g. 0.0845 = 8.45%)
-----------------------------------------------------------------
{
  "report_title": str, "as_of": "Sep-2026", "prior_month": "Mar-2026",
  "benchmark_name": str, "source_note": str, "data_status": str,
  "funds": [ 2 to 4 of {
      "short": "Fund A", "name": str, "category": str,
      "mandate_large_min": 0.80, "aum_cr": float, "expense": 0.0150,
      "cat_expense": 0.0140, "turnover": 0.45, "cash": 0.02,
      "investment": 400000, "manager": str,
      "holdings":      [ {"stock","sector","mcap","weight"} ... ],   # current month
      "prev_holdings": [ {"stock","sector","mcap","weight"} ... ]    # prior month
  } ],
  "benchmark": [ {"stock","sector","mcap","weight"} ... ],           # sums to 1.0
  "lineage": [ {"item","source","page_table","extracted_on","method","confidence"} ... ]
}
mcap must be one of "Large", "Mid", "Small".
"""
import argparse
import json
import random
import re
from itertools import combinations

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

# ----------------------------------------------------------------------------
# Styling helpers
# ----------------------------------------------------------------------------
NAVY, BLUE, GREEN, GREY = "1F3864", "0000FF", "008000", "808080"
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_SUB = PatternFill("solid", fgColor="D9E1F2")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FILL_NOTE = PatternFill("solid", fgColor="F2F2F2")
FILL_WARN = PatternFill("solid", fgColor="FCE4D6")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

PCT = '0.0%;(0.0%);"-"'
PCT2 = '0.00%;(0.00%);"-"'
NUM0 = '#,##0;(#,##0);"-"'
NUM1 = '#,##0.0;(#,##0.0);"-"'
RUPEE = '"₹"#,##0;("₹"#,##0);"-"'
RATIO = '0.00"x"'

LINK_RE = re.compile(r"^=('?[A-Za-z0-9_ ]+'?)!\$?[A-Z]{1,3}\$?\d+$")


def put(ws, ref, v, fmt=None, kind="calc", bold=False, fill=None, h=None,
        wrap=False, size=10, italic=False, color=None, border=False, v_al="center"):
    c = ws[ref]
    c.value = v
    if color is None:
        if kind == "input":
            color = BLUE
        elif isinstance(v, str) and LINK_RE.match(v):
            color = GREEN
        else:
            color = "000000"
    c.font = Font(name="Arial", size=size, bold=bold, italic=italic, color=color)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    c.alignment = Alignment(horizontal=h, vertical=v_al, wrap_text=wrap)
    if border:
        c.border = BOX
    return c


def rc(ws, r, c, v, **kw):
    return put(ws, f"{L(c)}{r}", v, **kw)


def hdr(ws, r, c, text, wrap=True):
    return put(ws, f"{L(c)}{r}", text, bold=True, fill=FILL_HDR, color="FFFFFF",
               h="center", wrap=wrap, border=True)


def title(ws, text, sub=None):
    put(ws, "A1", text, bold=True, size=14, color=NAVY)
    if sub:
        put(ws, "A2", sub, italic=True, color=GREY, size=9)


def section(ws, r, text, c1=1, c2=8):
    for c in range(c1, c2 + 1):
        ws.cell(row=r, column=c).fill = FILL_SUB
    put(ws, f"{L(c1)}{r}", text, bold=True, fill=FILL_SUB, color=NAVY)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def status_cf(ws, rng):
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"RED"'],
        fill=PatternFill("solid", bgColor="F8CBAD"), font=Font(name="Arial", bold=True, color="9C0006")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"AMBER"'],
        fill=PatternFill("solid", bgColor="FFE699"), font=Font(name="Arial", bold=True, color="7F6000")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"OK"'],
        fill=PatternFill("solid", bgColor="C6EFCE"), font=Font(name="Arial", bold=True, color="006100")))


def passfail_cf(ws, rng):
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"FAIL"'],
        fill=PatternFill("solid", bgColor="F8CBAD"), font=Font(name="Arial", bold=True, color="9C0006")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"PASS"'],
        fill=PatternFill("solid", bgColor="C6EFCE"), font=Font(name="Arial", bold=True, color="006100")))


def bar_chart(ws, title_, data_refs, cats, anchor, w=16, h=8, fmt=None, horizontal=False):
    ch = BarChart()
    ch.type = "bar" if horizontal else "col"
    ch.title = title_
    ch.style = 10
    for ref in data_refs:
        ch.add_data(ref, titles_from_data=True)
    ch.set_categories(cats)
    ch.width, ch.height = w, h
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    if fmt:
        ch.y_axis.number_format = fmt
    ws.add_chart(ch, anchor)


# ----------------------------------------------------------------------------
# Workbook builder
# ----------------------------------------------------------------------------
def build(cfg, path):
    funds = cfg["funds"]
    n = len(funds)
    assert 2 <= n <= 4, "Compare between 2 and 4 funds"
    pairs = list(combinations(range(n), 2))
    P = len(pairs)

    # ---- universe of stocks -------------------------------------------------
    master = {}

    def reg(h):
        key = (h["sector"], h["mcap"])
        if master.setdefault(h["stock"], key) != key:
            raise ValueError(f"Inconsistent sector/mcap for {h['stock']}")

    benchw = {h["stock"]: h["weight"] for h in cfg["benchmark"]}
    for f in funds:
        for h in f["holdings"] + f["prev_holdings"]:
            reg(h)
    for h in cfg["benchmark"]:
        reg(h)
    stocks = sorted(master, key=lambda s: (-benchw.get(s, 0), s))
    sectors = sorted({v[0] for v in master.values()})
    N, S = len(stocks), len(sectors)
    assert N >= 10, "Need at least 10 distinct stocks"

    wb = Workbook()
    wb._named_styles["Normal"].font = Font(name="Arial", size=10)
    names = ["Summary", "Inputs", "Holdings", "Holdings_Prev", "Benchmark", "Stock_Matrix",
             "Overlap", "Concentration", "Sector_Mcap", "Drift", "Red_Flags",
             "Combined_Portfolio", "Validation", "Methodology"]
    wb.active.title = names[0]
    for nm in names[1:]:
        wb.create_sheet(nm)
    S_ = {nm: wb[nm] for nm in names}
    tabs = {"Summary": "1F3864", "Inputs": "FFC000", "Holdings": "FFC000", "Holdings_Prev": "FFC000",
            "Benchmark": "FFC000", "Validation": "70AD47", "Methodology": "7F7F7F"}
    for nm, col in tabs.items():
        S_[nm].sheet_properties.tabColor = col

    # =========================================================================
    # INPUTS
    # =========================================================================
    ws = S_["Inputs"]
    title(ws, "Inputs & Assumptions",
          "Blue text on yellow = cells you can edit. Everything else in the workbook is formula-driven.")
    meta = [("Report title", cfg["report_title"]), ("Current disclosure month", cfg["as_of"]),
            ("Prior month compared", cfg["prior_month"]), ("Benchmark", cfg["benchmark_name"]),
            ("Data source", cfg["source_note"]), ("Data status", cfg["data_status"])]
    for k, (lab, val) in enumerate(meta):
        put(ws, f"A{4 + k}", lab, bold=True)
        put(ws, f"B{4 + k}", val, kind="input", fill=FILL_INPUT, border=True)
        ws.merge_cells(f"B{4 + k}:F{4 + k}")
    FR0 = 12  # first fund row
    heads = ["Short name", "Fund name", "Category", "Mandate: min large-cap % of NAV", "AUM (₹ Cr)",
             "Expense ratio", "Category avg expense", "Turnover ratio", "Cash % of NAV",
             "Investment (₹)", "Allocation %", "Fund manager"]
    for j, t in enumerate(heads):
        hdr(ws, FR0 - 1, 1 + j, t)
    ws.row_dimensions[FR0 - 1].height = 42
    for i, f in enumerate(funds):
        r = FR0 + i
        vals = [f["short"], f["name"], f["category"], f["mandate_large_min"], f["aum_cr"], f["expense"],
                f["cat_expense"], f["turnover"], f["cash"], f["investment"]]
        fm = [None, None, None, PCT, NUM0, PCT2, PCT2, PCT, PCT, RUPEE]
        for j, (v, fmt_) in enumerate(zip(vals, fm)):
            rc(ws, r, 1 + j, "Unavailable" if v is None else v, kind="input", fmt=fmt_, fill=FILL_INPUT, border=True)
        rc(ws, r, 11, f"=IF(SUM($J${FR0}:$J${FR0 + n - 1})=0,0,J{r}/SUM($J${FR0}:$J${FR0 + n - 1}))",
           fmt=PCT, border=True)
        rc(ws, r, 12, f.get("manager", ""), kind="input", fill=FILL_INPUT, border=True)
    TOTR = FR0 + n
    put(ws, f"A{TOTR}", "Total", bold=True)
    put(ws, f"J{TOTR}", f"=SUM(J{FR0}:J{TOTR - 1})", fmt=RUPEE, bold=True)
    put(ws, f"K{TOTR}", f"=SUM(K{FR0}:K{TOTR - 1})", fmt=PCT, bold=True)
    put(ws, f"A{TOTR + 1}", "Tip: set a fund's Investment to 0 to see the combined portfolio without it "
        "(Combined_Portfolio tab and Summary update; pairwise overlap is unchanged).", italic=True, color=GREY, size=9)

    THR = [
        ("hi_ov", "High overlap (pair)", 0.50, PCT, "Pairwise weighted overlap at/above this is rated HIGH", 15),
        ("mod_ov", "Moderate overlap (pair)", 0.30, PCT, "At/above this (but below HIGH) is rated MODERATE", None),
        ("act", "Minimum active share", 0.50, PCT, "Below this the fund looks like a 'closet index fund'", 12),
        ("mand", "Large-cap % ÷ mandate floor (min)", 1.00, RATIO, "Below 1.00x means the fund holds less large-cap than its mandate requires", 15),
        ("single", "Max single-stock weight", 0.10, PCT, "Largest single holding above this is a concentration risk", 8),
        ("top10", "Max top-10 concentration", 0.65, PCT, "Top 10 holdings above this share of NAV", 8),
        ("sector", "Max single-sector weight", 0.35, PCT, "Largest sector above this share of NAV", 8),
        ("effn", "Min effective number of stocks", 15, NUM1, "1/HHI. Low = money is bunched into few names", 8),
        ("drift", "Max portfolio change vs prior month", 0.25, PCT, "Share of the portfolio reshuffled since the prior month", 8),
        ("expgap", "Max expense gap vs category avg", 0.0025, PCT2, "Fund's expense ratio minus category average", 6),
        ("cash", "Max cash", 0.05, PCT, "Cash above this can drag returns", 6),
        ("turn", "Max portfolio turnover", 1.00, PCT, "Turnover above this signals heavy trading", 6),
        ("amber", "Amber warning margin", 0.15, PCT, "Within this margin of a limit = AMBER", None),
        ("chg", "Min change to call a holding increased/reduced", 0.001, PCT2, "Smaller moves are shown as 'Unchanged'", None),
        ("tol", "Validation tolerance", 0.0001, '0.0000', "Rounding tolerance for the checks on the Validation tab", None),
    ]
    TH0 = TOTR + 5
    section(ws, TH0 - 2, "Thresholds used by the Red-Flag engine (edit to suit your policy)", 1, 12)
    for j, t in enumerate(["Parameter", "Threshold", "What it means", "", "Penalty points"]):
        if t:
            hdr(ws, TH0 - 1, 1 + j, t)
    ws.merge_cells(f"C{TH0 - 1}:D{TH0 - 1}")
    thr_cell, pen_cell = {}, {}
    for k, (key, lab, val, fmt_, expl, pen) in enumerate(THR):
        r = TH0 + k
        put(ws, f"A{r}", lab, border=True)
        put(ws, f"B{r}", val, kind="input", fmt=fmt_, fill=FILL_INPUT, border=True)
        put(ws, f"C{r}", expl, size=9, color=GREY)
        if pen is not None:
            put(ws, f"E{r}", pen, kind="input", fmt="0", fill=FILL_INPUT, border=True)
            pen_cell[key] = f"Inputs!$E${r}"
        thr_cell[key] = f"Inputs!$B${r}"
    r = TH0 + len(THR)
    put(ws, f"A{r}", "Total penalty points (should be 100)", bold=True)
    put(ws, f"E{r}", f"=SUM(E{TH0}:E{r - 1})", fmt="0", bold=True)
    widths(ws, {"A": 36, "B": 34, "C": 14, "D": 16, "E": 12, "F": 12, "G": 14, "H": 12, "I": 12,
                "J": 16, "K": 12, "L": 22})
    ws.freeze_panes = "A4"
    fname = lambda i: f"Inputs!$B${FR0 + i}"
    fshort = lambda i: f"Inputs!$A${FR0 + i}"

    # =========================================================================
    # HOLDINGS / HOLDINGS_PREV / BENCHMARK  (raw data = the only hard-codes)
    # =========================================================================
    def write_holdings(ws, key, label):
        title(ws, label, "Raw disclosed data (blue = input). Paste/replace rows to refresh; regenerate to rebuild the stock list.")
        for j, t in enumerate(["Fund", "Stock", "Sector", "Market-cap bucket", "Weight (% of NAV)"]):
            hdr(ws, 4, 1 + j, t)
        r = 5
        for f in funds:
            for h in sorted(f[key], key=lambda x: -x["weight"]):
                for j, v in enumerate([f["name"], h["stock"], h["sector"], h["mcap"], h["weight"]]):
                    rc(ws, r, 1 + j, v, kind="input", fmt=PCT2 if j == 4 else None)
                r += 1
        widths(ws, {"A": 30, "B": 30, "C": 22, "D": 18, "E": 18})
        ws.freeze_panes = "A5"
        ws.auto_filter.ref = f"A4:E{r - 1}"
        return r - 1

    H_LAST = write_holdings(S_["Holdings"], "holdings", "Holdings - current month")
    write_holdings(S_["Holdings_Prev"], "prev_holdings", "Holdings - prior month")

    ws = S_["Benchmark"]
    title(ws, "Benchmark constituents", "Used for active share. Weights should sum to 100%.")
    for j, t in enumerate(["Stock", "Sector", "Market-cap bucket", "Weight (% of index)"]):
        hdr(ws, 4, 1 + j, t)
    for k, h in enumerate(sorted(cfg["benchmark"], key=lambda x: -x["weight"])):
        for j, v in enumerate([h["stock"], h["sector"], h["mcap"], h["weight"]]):
            rc(ws, 5 + k, 1 + j, v, kind="input", fmt=PCT2 if j == 3 else None)
    widths(ws, {"A": 30, "B": 22, "C": 18, "D": 18})
    ws.freeze_panes = "A5"

    # =========================================================================
    # STOCK_MATRIX  (engine room)
    # =========================================================================
    SM, F0 = "Stock_Matrix", 5
    LS = F0 + N - 1
    TOT = LS + 2
    cur = [L(4 + i) for i in range(n)]
    prv = [L(4 + n + i) for i in range(n)]
    bch, cmb, cnt = L(4 + 2 * n), L(5 + 2 * n), L(6 + 2 * n)
    pc = [L(7 + 2 * n + k) for k in range(P)]
    ab = [L(7 + 2 * n + P + i) for i in range(n)]
    ch = [L(7 + 2 * n + P + n + i) for i in range(n)]
    ctr = L(7 + 4 * n + P)
    kov, kcb = L(8 + 4 * n + P), L(9 + 4 * n + P)
    Rg = lambda col: f"{SM}!${col}${F0}:${col}${LS}"
    Tt = lambda col: f"{SM}!${col}${TOT}"
    pidx = {pr: k for k, pr in enumerate(pairs)}
    allocrow = f"{SM}!${cur[0]}$2:${cur[-1]}$2"

    ws = S_[SM]
    title(ws, "Stock Matrix - one row per stock, weights by fund (the calculation engine)",
          "All cells are formulas pulling from Holdings / Holdings_Prev / Benchmark. Grey columns are helpers.")
    put(ws, "C2", "Allocation →", bold=True, h="right")
    put(ws, "C3", "Lookup name →", italic=True, color=GREY, h="right")
    for i in range(n):
        put(ws, f"{cur[i]}2", f"=Inputs!$K${FR0 + i}", fmt=PCT)
        put(ws, f"{cur[i]}3", f"={fname(i)}", size=8, color=GREY, wrap=True)
        put(ws, f"{prv[i]}3", f"={fname(i)}", size=8, color=GREY, wrap=True)
        hdr(ws, 4, 4 + i, f"={fshort(i)}")
        hdr(ws, 4, 4 + n + i, f'={fshort(i)}&" (prior)"')
        hdr(ws, 4, 7 + 2 * n + P + i, f'="|W| vs bench: "&{fshort(i)}')
        hdr(ws, 4, 7 + 2 * n + P + n + i, f'="Change: "&{fshort(i)}')
    put(ws, f"{bch}3", "=Inputs!$B$7", size=8, color=GREY, wrap=True)
    for c, t in [(1, "Stock"), (2, "Sector"), (3, "Cap bucket")]:
        hdr(ws, 4, c, t)
    hdr(ws, 4, 4 + 2 * n, "Benchmark")
    hdr(ws, 4, 5 + 2 * n, "Combined (by allocation)")
    hdr(ws, 4, 6 + 2 * n, "# included funds holding")
    for (i, j), k in pidx.items():
        hdr(ws, 4, 7 + 2 * n + k, f'={fshort(i)}&" × "&{fshort(j)}')
    hdr(ws, 4, 7 + 4 * n + P, "Overlap contribution (avg over pairs)")
    hdr(ws, 4, 8 + 4 * n + P, "key: overlap (helper)")
    hdr(ws, 4, 9 + 4 * n + P, "key: combined (helper)")
    ws.row_dimensions[4].height = 44
    ws.row_dimensions[3].height = 30
    for k, s in enumerate(stocks):
        r = F0 + k
        put(ws, f"A{r}", s, bold=True)
        put(ws, f"B{r}", master[s][0])
        put(ws, f"C{r}", master[s][1])
        for i in range(n):
            put(ws, f"{cur[i]}{r}",
                f"=SUMIFS(Holdings!$E$5:$E$2000,Holdings!$A$5:$A$2000,{cur[i]}$3,Holdings!$B$5:$B$2000,$A{r})", fmt=PCT2)
            put(ws, f"{prv[i]}{r}",
                f"=SUMIFS(Holdings_Prev!$E$5:$E$2000,Holdings_Prev!$A$5:$A$2000,{prv[i]}$3,Holdings_Prev!$B$5:$B$2000,$A{r})", fmt=PCT2)
            put(ws, f"{ab[i]}{r}", f"=ABS({cur[i]}{r}-${bch}{r})", fmt=PCT2, color=GREY)
            put(ws, f"{ch[i]}{r}", f"={cur[i]}{r}-{prv[i]}{r}", fmt=PCT2)
        put(ws, f"{bch}{r}", f"=SUMIFS(Benchmark!$D$5:$D$2000,Benchmark!$A$5:$A$2000,$A{r})", fmt=PCT2)
        put(ws, f"{cmb}{r}",
            f"=IF(SUM(${cur[0]}$2:${cur[-1]}$2)=0,0,SUMPRODUCT(${cur[0]}$2:${cur[-1]}$2,{cur[0]}{r}:{cur[-1]}{r})/SUM(${cur[0]}$2:${cur[-1]}$2))",
            fmt=PCT2)
        put(ws, f"{cnt}{r}",
            f"=SUMPRODUCT(--({cur[0]}{r}:{cur[-1]}{r}>0),--(${cur[0]}$2:${cur[-1]}$2>0))", fmt="0")
        for (i, j), kk in pidx.items():
            put(ws, f"{pc[kk]}{r}", f"=MIN({cur[i]}{r},{cur[j]}{r})", fmt=PCT2, color=GREY)
        put(ws, f"{ctr}{r}", f"=SUM({pc[0]}{r}:{pc[-1]}{r})/{P}", fmt=PCT2)
        put(ws, f"{kov}{r}", f"={ctr}{r}+ROW()*0.000000000001", fmt="0.000000", color=GREY)
        put(ws, f"{kcb}{r}", f"={cmb}{r}+ROW()*0.000000000001", fmt="0.000000", color=GREY)
    put(ws, f"A{TOT}", "TOTAL", bold=True)
    sumcols = cur + prv + [bch, cmb] + pc + ab + [ctr]
    for col in sumcols:
        put(ws, f"{col}{TOT}", f"=SUM({col}{F0}:{col}{LS})", fmt=PCT2, bold=True)
    for col in ch:
        put(ws, f"{col}{TOT}", f"=SUM({col}{F0}:{col}{LS})", fmt=PCT2, bold=True)
    put(ws, f"A{TOT + 1}", "Pair columns: min(weight in fund X, weight in fund Y) per stock; their TOTAL is the weighted overlap. "
        "Sector / cap bucket are stock master data taken from the holdings files.", italic=True, color=GREY, size=9)
    widths(ws, {"A": 30, "B": 20, "C": 11})
    for c in range(4, 10 + 4 * n + P):
        ws.column_dimensions[L(c)].width = 13
    ws.freeze_panes = "D5"

    # =========================================================================
    # OVERLAP
    # =========================================================================
    ws = S_["Overlap"]
    title(ws, "Portfolio Overlap Analysis",
          "Weighted overlap = Σ over common stocks of the SMALLER of the two weights. 100% = identical portfolios, 0% = nothing in common.")
    section(ws, 4, "A. Weighted overlap matrix (% of portfolio in the same stocks)", 1, 1 + n)
    for j in range(n):
        hdr(ws, 5, 2 + j, f"={fshort(j)}")
    hdr(ws, 5, 1, "")
    for i in range(n):
        put(ws, f"A{6 + i}", f"={fshort(i)}", bold=True, border=True)
        for j in range(n):
            if i == j:
                rc(ws, 6 + i, 2 + j, 1, fmt="0%", border=True, color=GREY)
            else:
                k = pidx[(min(i, j), max(i, j))]
                rc(ws, 6 + i, 2 + j, f"={Tt(pc[k])}", fmt="0.0%", border=True, h="center")
    mrng = f"B6:{L(1 + n)}{5 + n}"
    ws.conditional_formatting.add(mrng, ColorScaleRule(start_type="num", start_value=0, start_color="C6EFCE",
                                                       mid_type="num", mid_value=0.5, mid_color="FFEB9C",
                                                       end_type="num", end_value=1, end_color="F8696B"))
    c0 = n + 8
    section(ws, c0 - 1, "B. Number of common stocks (diagonal = holdings count)", 1, 1 + n)
    for j in range(n):
        hdr(ws, c0, 2 + j, f"={fshort(j)}")
    hdr(ws, c0, 1, "")
    for i in range(n):
        put(ws, f"A{c0 + 1 + i}", f"={fshort(i)}", bold=True, border=True)
        for j in range(n):
            f_ = (f'=COUNTIF({Rg(cur[i])},">0")' if i == j
                  else f'=COUNTIFS({Rg(cur[i])},">0",{Rg(cur[j])},">0")')
            rc(ws, c0 + 1 + i, 2 + j, f_, fmt="0", border=True, h="center")
    pt = c0 + n + 4
    section(ws, pt - 1, "C. Pair ranking", 1, 6)
    for j, t in enumerate(["Pair", "Weighted overlap", "Common stocks", "Common as % of smaller fund's holdings", "Rating"]):
        hdr(ws, pt, 1 + j, t)
    ws.row_dimensions[pt].height = 42
    PT1, PT2 = pt + 1, pt + P
    for (i, j), k in pidx.items():
        r = PT1 + k
        put(ws, f"A{r}", f'={fshort(i)}&" × "&{fshort(j)}', border=True)
        put(ws, f"B{r}", f"={Tt(pc[k])}", fmt="0.0%", border=True)
        put(ws, f"C{r}", f'=COUNTIFS({Rg(cur[i])},">0",{Rg(cur[j])},">0")', fmt="0", border=True)
        put(ws, f"D{r}", f'=C{r}/MIN(COUNTIF({Rg(cur[i])},">0"),COUNTIF({Rg(cur[j])},">0"))', fmt="0%", border=True)
        put(ws, f"E{r}", f'=IF(B{r}>={thr_cell["hi_ov"]},"HIGH",IF(B{r}>={thr_cell["mod_ov"]},"MODERATE","LOW"))',
            border=True, h="center")
    ws.conditional_formatting.add(f"E{PT1}:E{PT2}", CellIsRule(operator="equal", formula=['"HIGH"'],
        fill=PatternFill("solid", bgColor="F8CBAD")))
    ws.conditional_formatting.add(f"E{PT1}:E{PT2}", CellIsRule(operator="equal", formula=['"MODERATE"'],
        fill=PatternFill("solid", bgColor="FFE699")))
    ws.conditional_formatting.add(f"E{PT1}:E{PT2}", CellIsRule(operator="equal", formula=['"LOW"'],
        fill=PatternFill("solid", bgColor="C6EFCE")))
    st = PT2 + 3
    section(ws, st - 1, "D. Headline overlap statistics", 1, 3)
    pa, pb, pcn = f"$A${PT1}:$A${PT2}", f"$B${PT1}:$B${PT2}", f"$C${PT1}:$C${PT2}"
    stats = [
        ("Average pair overlap", f"=AVERAGE({pb})", "0.0%"),
        ("Highest pair overlap", f"=MAX({pb})", "0.0%"),
        ("Lowest pair overlap", f"=MIN({pb})", "0.0%"),
        ("Most similar pair", f"=INDEX({pa},MATCH(MAX({pb}),{pb},0))", None),
        ("Common stocks in most similar pair", f"=INDEX({pcn},MATCH(MAX({pb}),{pb},0))", "0"),
        ("Least similar pair", f"=INDEX({pa},MATCH(MIN({pb}),{pb},0))", None),
        ("Pairs rated HIGH", f'=COUNTIF($E${PT1}:$E${PT2},"HIGH")', "0"),
    ]
    OV = {}
    for k, (lab, f_, fmt_) in enumerate(stats):
        put(ws, f"A{st + k}", lab, border=True)
        put(ws, f"B{st + k}", f_, fmt=fmt_, bold=True, border=True)
        OV[k] = f"Overlap!$B${st + k}"
    OV_AVG, OV_MAX, OV_MIN, OV_MAXP, OV_MAXC, OV_MINP, OV_NHI = (OV[k] for k in range(7))
    tp = st + len(stats) + 3
    section(ws, tp - 1, "E. Top 10 stocks driving the overlap", 1, 5)
    for j, t in enumerate(["Rank", "Stock", "Overlap contribution", "# funds holding", "Combined weight"]):
        hdr(ws, tp, 1 + j, t)
    for k in range(10):
        r = tp + 1 + k
        m = f"MATCH(LARGE({Rg(kov)},$A{r}),{Rg(kov)},0)"
        put(ws, f"A{r}", k + 1, fmt="0", border=True, h="center", color="000000")
        put(ws, f"B{r}", f"=INDEX({Rg('A')},{m})", border=True)
        put(ws, f"C{r}", f"=INDEX({Rg(ctr)},{m})", fmt=PCT2, border=True)
        put(ws, f"D{r}", f"=INDEX({Rg(cnt)},{m})", fmt="0", border=True, h="center")
        put(ws, f"E{r}", f"=INDEX({Rg(cmb)},{m})", fmt=PCT2, border=True)
    bar_chart(ws, "Top 10 overlap drivers", [Reference(ws, min_col=3, min_row=tp, max_row=tp + 10)],
              Reference(ws, min_col=2, min_row=tp + 1, max_row=tp + 10), f"H{4}", w=17, h=9, fmt="0.0%", horizontal=True)
    widths(ws, {"A": 34, "B": 22, "C": 18, "D": 20, "E": 18, "F": 14})
    OVROW = lambda i: f"Overlap!$B${6 + i}:${L(1 + n)}${6 + i}"

    # =========================================================================
    # SECTOR_MCAP
    # =========================================================================
    SC = "Sector_Mcap"
    ws = S_[SC]
    title(ws, "Sector & Market-Cap Analysis", "Sector weights by fund vs benchmark, active bets, drift since prior month, sector-level overlap.")
    s1, sL = 5, 4 + S
    sc_cur = [L(2 + i) for i in range(n)]
    sc_prv = [L(2 + n + i) for i in range(n)]
    sc_b, sc_c = L(2 + 2 * n), L(3 + 2 * n)
    sc_act = [L(4 + 2 * n + i) for i in range(n)]
    sc_chg = [L(4 + 3 * n + i) for i in range(n)]
    hdr(ws, 4, 1, "Sector")
    for i in range(n):
        hdr(ws, 4, 2 + i, f"={fshort(i)}")
        hdr(ws, 4, 2 + n + i, f'={fshort(i)}&" (prior)"')
        hdr(ws, 4, 4 + 2 * n + i, f'="Active bet: "&{fshort(i)}')
        hdr(ws, 4, 4 + 3 * n + i, f'="Change: "&{fshort(i)}')
    hdr(ws, 4, 2 + 2 * n, "Benchmark")
    hdr(ws, 4, 3 + 2 * n, "Combined")
    ws.row_dimensions[4].height = 40
    for k, sec in enumerate(sectors):
        r = s1 + k
        put(ws, f"A{r}", sec, bold=True, border=True)
        for i in range(n):
            put(ws, f"{sc_cur[i]}{r}", f"=SUMIFS({Rg(cur[i])},{Rg('B')},$A{r})", fmt=PCT, border=True)
            put(ws, f"{sc_prv[i]}{r}", f"=SUMIFS({Rg(prv[i])},{Rg('B')},$A{r})", fmt=PCT, border=True)
            put(ws, f"{sc_act[i]}{r}", f"={sc_cur[i]}{r}-${sc_b}{r}", fmt=PCT, border=True)
            put(ws, f"{sc_chg[i]}{r}", f"={sc_cur[i]}{r}-{sc_prv[i]}{r}", fmt=PCT, border=True)
        put(ws, f"{sc_b}{r}", f"=SUMIFS({Rg(bch)},{Rg('B')},$A{r})", fmt=PCT, border=True)
        put(ws, f"{sc_c}{r}", f"=SUMIFS({Rg(cmb)},{Rg('B')},$A{r})", fmt=PCT, border=True)
    put(ws, f"A{sL + 1}", "TOTAL", bold=True)
    for col in sc_cur + sc_prv + [sc_b, sc_c]:
        put(ws, f"{col}{sL + 1}", f"=SUM({col}{s1}:{col}{sL})", fmt=PCT, bold=True)
    for col in sc_act + sc_chg:
        ws.conditional_formatting.add(f"{col}{s1}:{col}{sL}", ColorScaleRule(
            start_type="min", start_color="F8696B", mid_type="num", mid_value=0, mid_color="FFFFFF",
            end_type="max", end_color="63BE7B"))
    so = sL + 5
    section(ws, so - 1, "Sector-level overlap (same sectors, ignoring which stock)", 1, 1 + n)
    hdr(ws, so, 1, "")
    for j in range(n):
        hdr(ws, so, 2 + j, f"={fshort(j)}")
    for i in range(n):
        put(ws, f"A{so + 1 + i}", f"={fshort(i)}", bold=True, border=True)
        for j in range(n):
            if i == j:
                rc(ws, so + 1 + i, 2 + j, 1, fmt="0%", border=True, color=GREY)
            else:
                a = f"${sc_cur[i]}${s1}:${sc_cur[i]}${sL}"
                b = f"${sc_cur[j]}${s1}:${sc_cur[j]}${sL}"
                rc(ws, so + 1 + i, 2 + j, f"=SUMPRODUCT(({a}+{b}-ABS({a}-{b}))/2)", fmt="0.0%", border=True, h="center")
    ws.conditional_formatting.add(f"B{so + 1}:{L(1 + n)}{so + n}", ColorScaleRule(
        start_type="num", start_value=0, start_color="C6EFCE", mid_type="num", mid_value=0.5,
        mid_color="FFEB9C", end_type="num", end_value=1, end_color="F8696B"))
    bar_chart(ws, "Sector weights: funds vs benchmark",
              [Reference(ws, min_col=2, max_col=1 + n, min_row=4, max_row=sL),
               Reference(ws, min_col=2 + 2 * n, min_row=4, max_row=sL)],
              Reference(ws, min_col=1, min_row=s1, max_row=sL), f"A{so + n + 4}", w=26, h=9, fmt="0%")
    widths(ws, {"A": 26})
    for c in range(2, 4 + 4 * n):
        ws.column_dimensions[L(c)].width = 13
    ws.freeze_panes = "B5"
    SCR = lambda col: f"{SC}!${col}${s1}:${col}${sL}"

    # =========================================================================
    # CONCENTRATION
    # =========================================================================
    CO = "Concentration"
    ws = S_[CO]
    title(ws, "Concentration, Style, Active Share & Drift - per fund",
          "Each column is a fund; the last column treats all funds together as one combined portfolio.")
    order = ["count", "eqexp", "cash", "lname", "lw", "top5", "top10", "hhi", "effn", "large", "mid", "small",
             "mandmin", "mandratio", "tsec", "tsecw", "active", "drift", "new", "exit", "expense", "expgap", "turn"]
    labels = {
        "count": ("Number of stock holdings", "0"), "eqexp": ("Equity exposure (% of NAV)", PCT),
        "cash": ("Cash (% of NAV)", PCT), "lname": ("Largest holding", None), "lw": ("Largest holding weight", PCT),
        "top5": ("Top-5 concentration", PCT), "top10": ("Top-10 concentration", PCT),
        "hhi": ("HHI (0 = spread out, 1 = one stock)", "0.000"), "effn": ("Effective number of stocks (1/HHI)", NUM1),
        "large": ("Large-cap % of NAV", PCT), "mid": ("Mid-cap % of NAV", PCT), "small": ("Small-cap % of NAV", PCT),
        "mandmin": ("Mandate: minimum large-cap %", PCT), "mandratio": ("Large-cap ÷ mandate floor", RATIO),
        "tsec": ("Largest sector", None), "tsecw": ("Largest sector weight", PCT),
        "active": ("Active share vs benchmark", PCT), "drift": ("Portfolio change vs prior month", PCT),
        "new": ("New entries since prior month", "0"), "exit": ("Exits since prior month", "0"),
        "expense": ("Expense ratio", PCT2), "expgap": ("Expense gap vs category avg", PCT2),
        "turn": ("Portfolio turnover", PCT),
    }
    CR = {k: 5 + j for j, k in enumerate(order)}
    hdr(ws, 4, 1, "Metric")
    for i in range(n):
        hdr(ws, 4, 2 + i, f"={fname(i)}")
    hdr(ws, 4, 2 + n, "Combined portfolio")
    ws.row_dimensions[4].height = 46
    for k in order:
        put(ws, f"A{CR[k]}", labels[k][0], bold=True, border=True)
    for i in range(n):
        c = L(2 + i)
        w_ = Rg(cur[i])
        FR = FR0 + i
        fm = {
            "count": f'=COUNTIF({w_},">0")', "eqexp": f"=SUM({w_})",
            "cash": f'=IF(ISNUMBER(Inputs!$I${FR}),Inputs!$I${FR},"Unavailable")',
            "lname": f"=INDEX({Rg('A')},MATCH(MAX({w_}),{w_},0))", "lw": f"=MAX({w_})",
            "top5": f"=SUMPRODUCT(LARGE({w_},{{1,2,3,4,5}}))", "top10": f"=SUMPRODUCT(LARGE({w_},{{1,2,3,4,5,6,7,8,9,10}}))",
            "hhi": f"=SUMPRODUCT({w_},{w_})/SUM({w_})^2", "effn": f"=1/{c}{CR['hhi']}",
            "large": f'=SUMIFS({w_},{Rg("C")},"Large")', "mid": f'=SUMIFS({w_},{Rg("C")},"Mid")',
            "small": f'=SUMIFS({w_},{Rg("C")},"Small")', "mandmin": f"=Inputs!$D${FR}",
            "mandratio": f"=IF({c}{CR['mandmin']}=0,99,{c}{CR['large']}/{c}{CR['mandmin']})",
            "tsec": f"=INDEX({SCR('A')},MATCH(MAX({SCR(sc_cur[i])}),{SCR(sc_cur[i])},0))",
            "tsecw": f"=MAX({SCR(sc_cur[i])})",
            "active": f'=IF(ISNUMBER(Inputs!$I${FR}),0.5*({Tt(ab[i])}+{c}{CR["cash"]}),"Unavailable")',
            "drift": f"=0.5*SUMPRODUCT(ABS({Rg(ch[i])}))",
            "new": f"=SUMPRODUCT(({Rg(cur[i])}>0)*({Rg(prv[i])}=0))",
            "exit": f"=SUMPRODUCT(({Rg(cur[i])}=0)*({Rg(prv[i])}>0))",
            "expense": f'=IF(ISNUMBER(Inputs!$F${FR}),Inputs!$F${FR},"Unavailable")',
            "expgap": f'=IF(AND(ISNUMBER(Inputs!$F${FR}),ISNUMBER(Inputs!$G${FR})),Inputs!$F${FR}-Inputs!$G${FR},"Unavailable")',
            "turn": f'=IF(ISNUMBER(Inputs!$H${FR}),Inputs!$H${FR},"Unavailable")',
        }
        for k in order:
            put(ws, f"{c}{CR[k]}", fm[k], fmt=labels[k][1], border=True, h="center")
    cc = L(2 + n)
    w_ = Rg(cmb)
    comb = {
        "count": f'=COUNTIF({w_},">0")', "eqexp": f"=SUM({w_})",
        "lname": f"=INDEX({Rg('A')},MATCH(MAX({w_}),{w_},0))", "lw": f"=MAX({w_})",
        "top5": f"=SUMPRODUCT(LARGE({w_},{{1,2,3,4,5}}))/SUM({w_})",
        "top10": f"=SUMPRODUCT(LARGE({w_},{{1,2,3,4,5,6,7,8,9,10}}))/SUM({w_})",
        "hhi": f"=SUMPRODUCT({w_},{w_})/SUM({w_})^2", "effn": f"={cc}{CR['hhi']}^-1",
        "large": f'=SUMIFS({w_},{Rg("C")},"Large")', "mid": f'=SUMIFS({w_},{Rg("C")},"Mid")',
        "small": f'=SUMIFS({w_},{Rg("C")},"Small")',
        "tsec": f"=INDEX({SCR('A')},MATCH(MAX({SCR(sc_c)}),{SCR(sc_c)},0))", "tsecw": f"=MAX({SCR(sc_c)})",
    }
    for k in order:
        put(ws, f"{cc}{CR[k]}", comb.get(k, "-"), fmt=labels[k][1], border=True, h="center",
            fill=FILL_NOTE if k not in comb else None)
    widths(ws, {"A": 40})
    for c in range(2, 3 + n):
        ws.column_dimensions[L(c)].width = 22
    ws.freeze_panes = "B5"
    CO_REF = lambda k, i: f"{CO}!${L(2 + i)}${CR[k]}"

    # =========================================================================
    # DRIFT
    # =========================================================================
    DR = "Drift"
    ws = S_[DR]
    title(ws, "Holding-level Drift - what changed since the prior month",
          "Pick a fund in the yellow cell. Everything below recalculates.")
    put(ws, "A3", "Select fund ▶", bold=True)
    put(ws, "B3", funds[0]["name"], kind="input", fill=FILL_INPUT, border=True)
    ws.merge_cells("B3:E3")
    dv = DataValidation(type="list", formula1=f"=Inputs!$B${FR0}:$B${FR0 + n - 1}", allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B3")
    put(ws, "A4", "Fund number", color=GREY)
    put(ws, "B4", f"=MATCH(B3,Inputs!$B${FR0}:$B${FR0 + n - 1},0)", fmt="0", color=GREY)
    dstat = [("Portfolio change vs prior month", f"=INDEX({CO}!$B${CR['drift']}:${L(1 + n)}${CR['drift']},$B$4)", PCT),
             ("New entries", f"=INDEX({CO}!$B${CR['new']}:${L(1 + n)}${CR['new']},$B$4)", "0"),
             ("Exits", f"=INDEX({CO}!$B${CR['exit']}:${L(1 + n)}${CR['exit']},$B$4)", "0")]
    for k, (lab, f_, fmt_) in enumerate(dstat):
        put(ws, f"A{5 + k}", lab, bold=True, border=True)
        put(ws, f"B{5 + k}", f_, fmt=fmt_, bold=True, border=True)
    D0 = 10
    for j, t in enumerate(["Stock", "Sector", "Current", "Prior", "Change", "Status", "key (helper)"]):
        hdr(ws, D0, 1 + j, t)
    for k in range(N):
        r, sr = D0 + 1 + k, F0 + k
        put(ws, f"A{r}", f"={SM}!$A${sr}", color="000000")
        put(ws, f"B{r}", f"={SM}!$B${sr}", color="000000")
        put(ws, f"C{r}", f"=INDEX({SM}!${cur[0]}{sr}:${cur[-1]}{sr},$B$4)", fmt=PCT2)
        put(ws, f"D{r}", f"=INDEX({SM}!${prv[0]}{sr}:${prv[-1]}{sr},$B$4)", fmt=PCT2)
        put(ws, f"E{r}", f"=C{r}-D{r}", fmt=PCT2)
        put(ws, f"F{r}", f'=IF(AND(C{r}=0,D{r}=0),"-",IF(AND(D{r}=0,C{r}>0),"New entry",IF(AND(C{r}=0,D{r}>0),"Exited",'
                          f'IF(E{r}>{thr_cell["chg"]},"Increased",IF(E{r}<-{thr_cell["chg"]},"Reduced","Unchanged")))))', h="center")
        put(ws, f"G{r}", f"=E{r}+ROW()*0.000000000001", fmt="0.000000", color=GREY)
    DL = D0 + N
    ws.conditional_formatting.add(f"E{D0 + 1}:E{DL}", ColorScaleRule(start_type="min", start_color="F8696B",
        mid_type="num", mid_value=0, mid_color="FFFFFF", end_type="max", end_color="63BE7B"))
    ws.conditional_formatting.add(f"F{D0 + 1}:F{DL}", CellIsRule(operator="equal", formula=['"New entry"'],
        fill=PatternFill("solid", bgColor="C6EFCE")))
    ws.conditional_formatting.add(f"F{D0 + 1}:F{DL}", CellIsRule(operator="equal", formula=['"Exited"'],
        fill=PatternFill("solid", bgColor="F8CBAD")))
    for j, t in enumerate(["Rank", "Biggest increases", "Change", "Biggest reductions", "Change"]):
        hdr(ws, D0, 9 + j, t)
    for k in range(5):
        r = D0 + 1 + k
        put(ws, f"I{r}", k + 1, fmt="0", h="center", border=True, color="000000")
        mi = f"MATCH(LARGE($G${D0 + 1}:$G${DL},$I{r}),$G${D0 + 1}:$G${DL},0)"
        md = f"MATCH(SMALL($G${D0 + 1}:$G${DL},$I{r}),$G${D0 + 1}:$G${DL},0)"
        put(ws, f"J{r}", f"=INDEX($A${D0 + 1}:$A${DL},{mi})", border=True)
        put(ws, f"K{r}", f"=INDEX($E${D0 + 1}:$E${DL},{mi})", fmt=PCT2, border=True)
        put(ws, f"L{r}", f"=INDEX($A${D0 + 1}:$A${DL},{md})", border=True)
        put(ws, f"M{r}", f"=INDEX($E${D0 + 1}:$E${DL},{md})", fmt=PCT2, border=True)
    widths(ws, {"A": 30, "B": 20, "C": 12, "D": 12, "E": 12, "F": 14, "G": 12, "H": 3, "I": 7, "J": 28, "K": 12, "L": 28, "M": 12})
    ws.freeze_panes = f"A{D0 + 1}"

    # =========================================================================
    # RED_FLAGS
    # =========================================================================
    RF = "Red_Flags"
    ws = S_[RF]
    title(ws, "Red-Flag Engine & Health Score",
          "Each rule compares a fund metric with a threshold (edit on Inputs). RED = breach, AMBER = within the warning margin.")
    rules = [
        ("hi_ov", "High overlap with another fund", "Highest weighted overlap between this fund and any other fund", "Max", PCT,
         lambda i: f"=LARGE({OVROW(i)},2)"),
        ("act", "Closet indexing", "Active share vs benchmark - low means it mostly copies the index", "Min", PCT,
         lambda i: f"={CO_REF('active', i)}"),
        ("mand", "Mandate drift (large-cap floor)", "Large-cap % divided by the mandate's minimum", "Min", RATIO,
         lambda i: f"={CO_REF('mandratio', i)}"),
        ("single", "Single-stock concentration", "Weight of the largest holding", "Max", PCT,
         lambda i: f"={CO_REF('lw', i)}"),
        ("top10", "Top-10 concentration", "Share of NAV in the ten largest holdings", "Max", PCT,
         lambda i: f"={CO_REF('top10', i)}"),
        ("sector", "Sector crowding", "Weight of the largest sector", "Max", PCT,
         lambda i: f"={CO_REF('tsecw', i)}"),
        ("effn", "Few effective stocks", "Effective number of stocks (1/HHI)", "Min", NUM1,
         lambda i: f"={CO_REF('effn', i)}"),
        ("drift", "Heavy reshuffling", "Share of portfolio changed since the prior month", "Max", PCT,
         lambda i: f"={CO_REF('drift', i)}"),
        ("expgap", "Expensive vs category", "Expense ratio minus category average", "Max", PCT2,
         lambda i: f"={CO_REF('expgap', i)}"),
        ("cash", "Cash drag", "Cash as % of NAV", "Max", PCT,
         lambda i: f"={CO_REF('cash', i)}"),
        ("turn", "High turnover", "Portfolio turnover ratio", "Max", PCT,
         lambda i: f"={CO_REF('turn', i)}"),
    ]
    R0 = 6
    RL = R0 + len(rules) - 1
    v0, s0 = 7, 8 + n
    for j, t in enumerate(["#", "Rule", "What it checks (plain English)", "Direction", "Threshold", "Penalty pts"]):
        hdr(ws, R0 - 1, 1 + j, t)
    for i in range(n):
        hdr(ws, R0 - 1, v0 + i, f'={fshort(i)}&" value"')
        hdr(ws, R0 - 1, s0 + i, f'={fshort(i)}&" status"')
    section(ws, R0 - 2, "Metric values", v0, v0 + n - 1)
    section(ws, R0 - 2, "Status", s0, s0 + n - 1)
    ws.row_dimensions[R0 - 1].height = 32
    for k, (key, nm, expl, dirn, fmt_, fn) in enumerate(rules):
        r = R0 + k
        put(ws, f"A{r}", k + 1, fmt="0", border=True, h="center", color="000000")
        put(ws, f"B{r}", nm, bold=True, border=True)
        put(ws, f"C{r}", expl, size=9, border=True, wrap=True)
        put(ws, f"D{r}", dirn, border=True, h="center", kind="input")
        put(ws, f"E{r}", f"={thr_cell[key]}", fmt=fmt_, border=True, h="center")
        put(ws, f"F{r}", f"={pen_cell[key]}", fmt="0", border=True, h="center")
        ws.row_dimensions[r].height = 28
        for i in range(n):
            vc = f"{L(v0 + i)}{r}"
            put(ws, vc, fn(i), fmt=fmt_, border=True, h="center")
            put(ws, f"{L(s0 + i)}{r}",
                f'=IF(NOT(ISNUMBER({vc})),"UNAVAILABLE",IF($D{r}="Max",IF({vc}>$E{r},"RED",IF({vc}>$E{r}*(1-{thr_cell["amber"]}),"AMBER","OK")),'
                f'IF({vc}<$E{r},"RED",IF({vc}<$E{r}*(1+{thr_cell["amber"]}),"AMBER","OK"))))',
                border=True, h="center")
    status_cf(ws, f"{L(s0)}{R0}:{L(s0 + n - 1)}{RL}")
    HS = RL + 2
    lab = {HS: "Health score (0-100)", HS + 1: "Number of RED flags", HS + 2: "Number of AMBER flags",
           HS + 3: "Rating", HS + 4: "First RED flag"}
    for r, t in lab.items():
        put(ws, f"B{r}", t, bold=True, border=True)
    for i in range(n):
        sc = f"{L(s0 + i)}${R0}:{L(s0 + i)}${RL}"
        c = L(v0 + i)
        put(ws, f"{c}{HS}", f'=MAX(0,100-SUMPRODUCT(({sc}="RED")*$F${R0}:$F${RL})-0.5*SUMPRODUCT(({sc}="AMBER")*$F${R0}:$F${RL}))',
            fmt="0", bold=True, border=True, h="center")
        put(ws, f"{c}{HS + 1}", f'=COUNTIF({sc},"RED")', fmt="0", border=True, h="center")
        put(ws, f"{c}{HS + 2}", f'=COUNTIF({sc},"AMBER")', fmt="0", border=True, h="center")
        put(ws, f"{c}{HS + 3}", f'=IF(COUNTIF({sc},"UNAVAILABLE")>0,"PARTIAL",IF({c}{HS}>=80,"HEALTHY",IF({c}{HS}>=60,"WATCH","CONCERN")))', bold=True, border=True, h="center")
        put(ws, f"{c}{HS + 4}", f'=IFERROR(INDEX($B${R0}:$B${RL},MATCH("RED",{sc},0)),"None")', border=True, h="center", wrap=True, size=9)
    ws.row_dimensions[HS + 4].height = 30
    ws.conditional_formatting.add(f"{L(v0)}{HS + 3}:{L(v0 + n - 1)}{HS + 3}", CellIsRule(operator="equal", formula=['"HEALTHY"'],
        fill=PatternFill("solid", bgColor="C6EFCE")))
    ws.conditional_formatting.add(f"{L(v0)}{HS + 3}:{L(v0 + n - 1)}{HS + 3}", CellIsRule(operator="equal", formula=['"WATCH"'],
        fill=PatternFill("solid", bgColor="FFE699")))
    ws.conditional_formatting.add(f"{L(v0)}{HS + 3}:{L(v0 + n - 1)}{HS + 3}", CellIsRule(operator="equal", formula=['"CONCERN"'],
        fill=PatternFill("solid", bgColor="F8CBAD")))
    put(ws, f"B{HS + 6}", "Score = 100 − penalty points for every RED rule − half the points for every AMBER rule (floor 0). "
        "Bands: 80+ HEALTHY, 60-79 WATCH, below 60 CONCERN.", italic=True, color=GREY, size=9)
    widths(ws, {"A": 5, "B": 32, "C": 44, "D": 11, "E": 11, "F": 10})
    for c in range(7, 9 + 2 * n):
        ws.column_dimensions[L(c)].width = 14
    ws.freeze_panes = f"C{R0}"
    RFV = lambda rule_idx, i: f"{RF}!${L(v0 + i)}${R0 + rule_idx}"
    RFH = lambda off, i: f"{RF}!${L(v0 + i)}${HS + off}"
    rf_status_rng = lambda rule_idx: f"{RF}!${L(s0)}${R0 + rule_idx}:${L(s0 + n - 1)}${R0 + rule_idx}"

    # =========================================================================
    # COMBINED_PORTFOLIO
    # =========================================================================
    CP = "Combined_Portfolio"
    ws = S_[CP]
    title(ws, "Combined (Look-through) Portfolio",
          "What you actually own once all funds are stacked together, weighted by the amounts on the Inputs tab.")
    incl = f'COUNTIF({allocrow},">0")'
    cpm = [
        ("Total invested", f"=Inputs!$J${TOTR}", RUPEE),
        ("Unique stocks owned", f'=COUNTIF({Rg(cmb)},">0")', "0"),
        ("Total stock positions (sum of each fund's holdings)", f"=SUM({Rg(cnt)})", "0"),
        ("Duplication ratio (positions ÷ unique stocks)", "=B6/B5", RATIO),
        ("Effective number of stocks (combined)", f"={Tt(cmb)}^2/SUMPRODUCT({Rg(cmb)},{Rg(cmb)})", NUM1),
        ("Top-10 share of combined money", f"=SUMPRODUCT(LARGE({Rg(cmb)},{{1,2,3,4,5,6,7,8,9,10}}))/{Tt(cmb)}", PCT),
        ("Stocks held by EVERY included fund", f"=COUNTIF({Rg(cnt)},{incl})", "0"),
        ("Share of money in stocks held by every fund", f"=SUMIF({Rg(cnt)},{incl},{Rg(cmb)})/{Tt(cmb)}", PCT),
        ("Share of money in stocks held by 2+ funds", f'=SUMIF({Rg(cnt)},">=2",{Rg(cmb)})/{Tt(cmb)}', PCT),
        ("Blended expense ratio", f'=IF(COUNT(Inputs!$F${FR0}:$F${FR0 + n - 1})<{n},"Unavailable",SUMPRODUCT(Inputs!$K${FR0}:$K${FR0 + n - 1},Inputs!$F${FR0}:$F${FR0 + n - 1}))', PCT2),
        ("Estimated annual fees (₹)", '=IF(NOT(ISNUMBER(B13)),"Unavailable",B4*B13)', RUPEE),
        ("Heuristic: fees paid on duplicated exposure (₹) = fees × average pair overlap", f'=IF(NOT(ISNUMBER(B14)),"Unavailable",B14*{OV_AVG})', RUPEE),
    ]
    for k, (lab_, f_, fmt_) in enumerate(cpm):
        put(ws, f"A{4 + k}", lab_, bold=True, border=True, wrap=True)
        put(ws, f"B{4 + k}", f_, fmt=fmt_, bold=True, border=True, h="center")
    put(ws, "A16", "The last line is an illustrative rule of thumb, not an accounting figure.", italic=True, color=GREY, size=9)
    CPR = 20
    section(ws, CPR - 1, "Top 15 look-through holdings", 1, 6)
    for j, t in enumerate(["Rank", "Stock", "Sector", "Combined weight", "₹ exposure", "# funds holding"]):
        hdr(ws, CPR, 1 + j, t)
    for k in range(15):
        r = CPR + 1 + k
        m = f"MATCH(LARGE({Rg(kcb)},$A{r}),{Rg(kcb)},0)"
        put(ws, f"A{r}", k + 1, fmt="0", h="center", border=True, color="000000")
        put(ws, f"B{r}", f"=INDEX({Rg('A')},{m})", border=True)
        put(ws, f"C{r}", f"=INDEX({Rg('B')},{m})", border=True)
        put(ws, f"D{r}", f"=INDEX({Rg(cmb)},{m})", fmt=PCT2, border=True)
        put(ws, f"E{r}", f"=D{r}*$B$4", fmt=RUPEE, border=True)
        put(ws, f"F{r}", f"=INDEX({Rg(cnt)},{m})", fmt="0", border=True, h="center")
    bar_chart(ws, "Top 15 look-through holdings", [Reference(ws, min_col=4, min_row=CPR, max_row=CPR + 15)],
              Reference(ws, min_col=2, min_row=CPR + 1, max_row=CPR + 15), "H4", w=18, h=11, fmt="0%", horizontal=True)
    widths(ws, {"A": 52, "B": 26, "C": 22, "D": 16, "E": 16, "F": 16})
    CP_TOP1N, CP_TOP1W = f"{CP}!$B${CPR + 1}", f"{CP}!$D${CPR + 1}"

    # =========================================================================
    # VALIDATION
    # =========================================================================
    VA = "Validation"
    ws = S_[VA]
    title(ws, "Validation - proof that the numbers are right",
          "Independent re-computations and integrity checks. Any FAIL means the data or extraction needs attention.")
    put(ws, "A3", "Overall status", bold=True)
    for j, t in enumerate(["Check", "Item", "Value A", "Value B", "Difference", "Status"]):
        hdr(ws, 5, 1 + j, t)
    tol = thr_cell["tol"]
    vrows = []

    def vcheck(check, item, a, b, fmt_, mode="eq"):
        vrows.append((check, item, a, b, fmt_, mode))

    for i in range(n):
        nm = f"={fname(i)}"
        vcheck("Equity + cash = 100% of NAV", nm,
               f'=IF(ISNUMBER(Inputs!$I${FR0 + i}),{Tt(cur[i])}+Inputs!$I${FR0 + i},"Unavailable")', "=1", PCT2)
        vcheck("Every current holding matched to the matrix", nm,
               f"=SUMIFS(Holdings!$E$5:$E$2000,Holdings!$A$5:$A$2000,{fname(i)})", f"={Tt(cur[i])}", PCT2)
        vcheck("Every prior holding matched to the matrix", nm,
               f"=SUMIFS(Holdings_Prev!$E$5:$E$2000,Holdings_Prev!$A$5:$A$2000,{fname(i)})", f"={Tt(prv[i])}", PCT2)
        vcheck("No duplicate stock lines for this fund", nm,
               f"=SUMPRODUCT((Holdings!$A$5:$A${H_LAST}={fname(i)})*(COUNTIFS(Holdings!$A$5:$A${H_LAST},Holdings!$A$5:$A${H_LAST},"
               f"Holdings!$B$5:$B${H_LAST},Holdings!$B$5:$B${H_LAST})>1))", "=0", "0")
        vcheck("Large + Mid + Small = equity exposure", nm,
               f"={CO_REF('large', i)}+{CO_REF('mid', i)}+{CO_REF('small', i)}", f"={CO_REF('eqexp', i)}", PCT2)
    vcheck("Benchmark weights sum to 100%", "Benchmark", "=SUM(Benchmark!$D$5:$D$2000)", "=1", PCT2)
    vcheck("Every benchmark stock matched to the matrix", "Benchmark", "=SUM(Benchmark!$D$5:$D$2000)", f"={Tt(bch)}", PCT2)
    vcheck("Sector table total = equity exposure (all funds)", "All funds",
           f"=SUM({SC}!${sc_cur[0]}${sL + 1}:${sc_cur[-1]}${sL + 1})", f"=SUM({Tt(cur[0])}:{Tt(cur[-1])})", PCT2)
    for (i, j), k in pidx.items():
        a = Rg(cur[i])
        b = Rg(cur[j])
        item = f'={fshort(i)}&" × "&{fshort(j)}'
        vcheck("Overlap: matrix method = independent SUMPRODUCT method", item, f"={Tt(pc[k])}",
               f"=SUMPRODUCT(({a}+{b}-ABS({a}-{b}))/2)", PCT2)
        vcheck("Overlap cannot exceed the smaller fund's equity exposure", item, f"={Tt(pc[k])}",
               f"=MIN({Tt(cur[i])},{Tt(cur[j])})", PCT2, mode="le")
    V0 = 6
    for k, (check, item, a, b, fmt_, mode) in enumerate(vrows):
        r = V0 + k
        put(ws, f"A{r}", check, border=True)
        put(ws, f"B{r}", item, border=True, wrap=True, color="000000")
        put(ws, f"C{r}", a, fmt=fmt_, border=True, h="center")
        put(ws, f"D{r}", b, fmt=fmt_, border=True, h="center")
        put(ws, f"E{r}", f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),C{r}-D{r},"Unavailable")', fmt='0.0000;-0.0000;0', border=True, h="center")
        cond = f"ABS(E{r})<={tol}" if mode == "eq" else f"E{r}<={tol}"
        put(ws, f"F{r}", f'=IF(NOT(ISNUMBER(E{r})),"UNAVAILABLE",IF({cond},"PASS","FAIL"))', border=True, h="center", bold=True)
    VL = V0 + len(vrows) - 1
    passfail_cf(ws, f"F{V0}:F{VL}")
    put(ws, "B3", f'=IF(COUNTIF(F{V0}:F{VL},"FAIL")>0,COUNTIF(F{V0}:F{VL},"FAIL")&" CHECK(S) FAILED",IF(COUNTIF(F{V0}:F{VL},"UNAVAILABLE")>0,"CHECKS PARTIALLY UNAVAILABLE","ALL "&COUNTA(F{V0}:F{VL})&" CHECKS PASSED"))',
        bold=True)
    widths(ws, {"A": 56, "B": 34, "C": 14, "D": 14, "E": 12, "F": 10})
    ws.freeze_panes = "A6"
    VAL_STATUS = f"{VA}!$B$3"

    # =========================================================================
    # SUMMARY
    # =========================================================================
    ws = S_["Summary"]
    put(ws, "A1", "=Inputs!$B$4", bold=True, size=16, color=NAVY)
    put(ws, "A2", '="Data as of "&Inputs!$B$5&"  |  compared with "&Inputs!$B$6&"  |  benchmark: "&Inputs!$B$7', italic=True, color=GREY)
    put(ws, "A3", "=Inputs!$B$9", bold=True, color="C00000", fill=FILL_WARN)
    ws.merge_cells("A3:L3")
    section(ws, 5, "IN PLAIN ENGLISH", 1, 12)
    am = f"{OV_AVG}"
    hot = f"INDEX({CO}!$B${CR['drift']}:${L(1 + n)}${CR['drift']},MATCH(MAX({CO}!$B${CR['drift']}:${L(1 + n)}${CR['drift']}),{CO}!$B${CR['drift']}:${L(1 + n)}${CR['drift']},0))"
    hotn = f"INDEX({CO}!$B$4:${L(1 + n)}$4,MATCH(MAX({CO}!$B${CR['drift']}:${L(1 + n)}${CR['drift']}),{CO}!$B${CR['drift']}:${L(1 + n)}${CR['drift']},0))"
    hsr = f"{RF}!${L(v0)}${HS}:${L(v0 + n - 1)}${HS}"
    low_name = f"INDEX({CO}!$B$4:${L(1 + n)}$4,MATCH(MIN({hsr}),{hsr},0))"
    low_rate = f"INDEX({RF}!${L(v0)}${HS + 3}:${L(v0 + n - 1)}${HS + 3},MATCH(MIN({hsr}),{hsr},0))"
    sec_top = f"INDEX({SCR('A')},MATCH(MAX({SCR(sc_c)}),{SCR(sc_c)},0))"
    lines = [
        f'="1. You compared "&{n}&" funds. On average, "&TEXT({am},"0%")&" of the money in any two of them sits in the very same stocks (weighted overlap)."',
        f'="2. The most alike pair is "&{OV_MAXP}&": "&TEXT({OV_MAX},"0%")&" overlap with "&{OV_MAXC}&" shared stocks. The most different pair is "&{OV_MINP}&" at "&TEXT({OV_MIN},"0%")&"."',
        (f'=IF({am}>={thr_cell["hi_ov"]},"3. Verdict: HIGH overlap. These funds behave much like one fund held several times, so you may be paying several fees for similar exposure.",'
         f'IF({am}>={thr_cell["mod_ov"]},"3. Verdict: MODERATE overlap. There is real diversification, but several of the same big stocks keep repeating.",'
         f'"3. Verdict: LOW overlap. The funds hold quite different stocks, so each one adds something different."))'),
        f'="4. On paper you hold "&{CP}!$B$6&" stock positions, but only "&{CP}!$B$5&" different stocks. Your combined money behaves like about "&TEXT({CP}!$B$8,"0")&" equally-sized stocks."',
        f'="5. The biggest shared bet is "&{CP_TOP1N}&" ("&TEXT({CP_TOP1W},"0.0%")&" of your combined money). The biggest sector is "&{sec_top}&" ("&TEXT(MAX({SCR(sc_c)}),"0%")&")."',
        f'="6. "&COUNTIF({rf_status_rng(1)},"RED")&" of "&{n}&" funds look like closet index funds (active share below "&TEXT({thr_cell["act"]},"0%")&" - they mostly copy the benchmark)."',
        f'="7. "&COUNTIF({rf_status_rng(2)},"RED")&" fund(s) hold less in large-cap stocks than their stated mandate requires."',
        f'="8. Since "&Inputs!$B$6&", the biggest portfolio change was in "&{hotn}&" ("&TEXT({hot},"0%")&" of its portfolio was reshuffled)."',
        f'="9. Health scores run from "&TEXT(MIN({hsr}),"0")&" to "&TEXT(MAX({hsr}),"0")&" out of 100. Weakest: "&{low_name}&" ("&{low_rate}&")."',
        f'="10. Data checks: "&{VAL_STATUS}&"."',
    ]
    for k, f_ in enumerate(lines):
        r = 6 + k
        put(ws, f"A{r}", f_, wrap=True, size=10, color="000000", v_al="center")
        ws.merge_cells(f"A{r}:L{r}")
        ws.row_dimensions[r].height = 30
    put(ws, "A16", "Educational analysis of publicly disclosed portfolios. Not investment advice. Disclosures are monthly snapshots and lag actual holdings.",
        italic=True, color=GREY, size=9)
    ws.merge_cells("A16:L16")

    SCT = 19
    section(ws, SCT - 1, "FUND SCORECARD", 1, 12)
    for j, t in enumerate(["Fund", "Category", "AUM (₹ Cr)", "Expense ratio", "Holdings", "Top-10 conc.",
                           "Effective stocks", "Active share", "Highest overlap with another fund", "RED flags",
                           "Health score", "Rating"]):
        hdr(ws, SCT, 1 + j, t)
    ws.row_dimensions[SCT].height = 44
    for i in range(n):
        r = SCT + 1 + i
        vals = [(f"={fname(i)}", None), (f"=Inputs!$C${FR0 + i}", None), (f"=Inputs!$E${FR0 + i}", NUM0),
                (f"=Inputs!$F${FR0 + i}", PCT2), (f"={CO_REF('count', i)}", "0"), (f"={CO_REF('top10', i)}", PCT),
                (f"={CO_REF('effn', i)}", NUM1), (f"={CO_REF('active', i)}", PCT), (f"={RFV(0, i)}", PCT),
                (f"={RFH(1, i)}", "0"), (f"={RFH(0, i)}", "0"), (f"={RFH(3, i)}", None)]
        for j, (f_, fmt_) in enumerate(vals):
            put(ws, f"{L(1 + j)}{r}", f_, fmt=fmt_, border=True, h="center" if j else None)
    ws.conditional_formatting.add(f"K{SCT + 1}:K{SCT + n}", ColorScaleRule(start_type="num", start_value=40, start_color="F8696B",
        mid_type="num", mid_value=70, mid_color="FFEB9C", end_type="num", end_value=100, end_color="63BE7B"))
    for txt, col in [("HEALTHY", "C6EFCE"), ("WATCH", "FFE699"), ("CONCERN", "F8CBAD")]:
        ws.conditional_formatting.add(f"L{SCT + 1}:L{SCT + n}", CellIsRule(operator="equal", formula=[f'"{txt}"'],
            fill=PatternFill("solid", bgColor=col)))
    HM = SCT + n + 3
    section(ws, HM - 1, "OVERLAP HEATMAP (weighted overlap between funds)", 1, 12)
    for j in range(n):
        hdr(ws, HM, 2 + j, f"={fshort(j)}")
    hdr(ws, HM, 1, "")
    for i in range(n):
        put(ws, f"A{HM + 1 + i}", f"={fshort(i)}", bold=True, border=True)
        for j in range(n):
            rc(ws, HM + 1 + i, 2 + j, f"=Overlap!${L(2 + j)}${6 + i}", fmt="0%", border=True, h="center", color="000000")
    ws.conditional_formatting.add(f"B{HM + 1}:{L(1 + n)}{HM + n}", ColorScaleRule(start_type="num", start_value=0,
        start_color="C6EFCE", mid_type="num", mid_value=0.5, mid_color="FFEB9C", end_type="num", end_value=1, end_color="F8696B"))
    AK = HM + n + 3
    section(ws, AK - 1, "WHAT TO LOOK AT NEXT", 1, 12)
    actions = [
        f'=IF({OV_NHI}>0,"• "&{OV_NHI}&" fund pair(s) rated HIGH overlap. Review "&{OV_MAXP}&": holding both adds little diversification.","• No fund pair breaches the HIGH-overlap threshold.")',
        f'=IF(COUNTIF({rf_status_rng(1)},"UNAVAILABLE")>0,"• Active-share check is unavailable for one or more funds because cash exposure is absent from the source workbook.",IF(COUNTIF({rf_status_rng(1)},"RED")>0,"• Closet-indexing warning: compare the fund fee with the cost of a plain index fund for the same exposure.","• All funds show meaningful active share."))',
        f'=IF(COUNTIF({rf_status_rng(2)},"RED")>0,"• Mandate warning: a fund holds less large-cap than its category rule. Check the latest factsheet for the reason.","• All funds meet their large-cap mandate floor.")',
        '="• Use the Drift tab to see which holdings each fund bought or sold since the prior month."',
        '="• Use the What-if idea: set a fund\'s Investment to 0 on the Inputs tab and watch the combined portfolio change."',
    ]
    for k, f_ in enumerate(actions):
        r = AK + k
        put(ws, f"A{r}", f_, wrap=True, color="000000")
        ws.merge_cells(f"A{r}:L{r}")
        ws.row_dimensions[r].height = 22
    widths(ws, {"A": 34, "B": 16, "C": 12, "D": 12, "E": 11, "F": 12, "G": 12, "H": 12, "I": 18, "J": 10, "K": 11, "L": 12})
    bar_chart(ws, "Health score by fund", [Reference(ws, min_col=11, min_row=SCT, max_row=SCT + n)],
              Reference(ws, min_col=1, min_row=SCT + 1, max_row=SCT + n), "N5", w=16, h=8)
    bar_chart(ws, "Top-10 concentration vs active share",
              [Reference(ws, min_col=6, min_row=SCT, max_row=SCT + n), Reference(ws, min_col=8, min_row=SCT, max_row=SCT + n)],
              Reference(ws, min_col=1, min_row=SCT + 1, max_row=SCT + n), "N22", w=16, h=8, fmt="0%")
    ws.sheet_view.showGridLines = False

    # =========================================================================
    # METHODOLOGY
    # =========================================================================
    ws = S_["Methodology"]
    title(ws, "Methodology, Definitions & Data Lineage")
    rows = [
        ("HOW TO USE", None),
        ("Edit only blue cells (Inputs, Holdings, Holdings_Prev, Benchmark, Drift!B3). All other tabs are formulas.", None),
        ("Colour code: blue = input, black = formula, green = link to another sheet, grey = helper.", None),
        ("DEFINITIONS (simple English)", None),
        ("Weighted overlap", "For every stock two funds both own, take the smaller of the two weights and add them up. 60% means 60% of each portfolio (at most) is the same money in the same stocks."),
        ("Effective number of stocks", "1 ÷ HHI. A fund with 50 holdings but most money in 10 names behaves like a ~20-stock fund."),
        ("HHI", "Sum of squared weights (after scaling to 100%). Higher = more concentrated."),
        ("Active share", "Half the sum of |fund weight − benchmark weight| across all stocks (+ cash). 0% = identical to the index, 100% = nothing in common. Below ~50-60% for an 'active' fund is a closet-indexing warning (Cremers & Petajisto, 2009)."),
        ("Portfolio change vs prior month", "Half the sum of |change in weight| across equity holdings. Measures how much of the portfolio was reshuffled."),
        ("Mandate ratio", "Large-cap % of NAV divided by the minimum the category requires. Below 1.00x = mandate breach."),
        ("Health score", "100 minus penalty points for each RED rule (half for AMBER). Penalties and thresholds are editable on Inputs."),
        ("LIMITATIONS", None),
        ("Monthly disclosures are snapshots and lag actual trading. Cash and derivatives are not part of the overlap. Market-cap buckets should follow the AMFI half-yearly list.", None),
        ("Sector and cap-bucket labels come from the holdings files. When uploading real factsheets, map every stock to a unique ISIN first so name variants do not break overlap.", None),
        ("Combined-portfolio figures use the allocation set on Inputs; pairwise overlap is independent of allocation.", None),
    ]
    r = 3
    for a, b in rows:
        if b is None and a.isupper():
            section(ws, r, a, 1, 2)
        elif b is None:
            put(ws, f"A{r}", a, wrap=True)
            ws.merge_cells(f"A{r}:B{r}")
            ws.row_dimensions[r].height = 30
        else:
            put(ws, f"A{r}", a, bold=True, border=True, v_al="top")
            put(ws, f"B{r}", b, wrap=True, border=True)
            ws.row_dimensions[r].height = 44
        r += 1
    r += 1
    section(ws, r, "DATA LINEAGE (fill in for every upload)", 1, 6)
    for j, t in enumerate(["Item", "Source file / URL", "Page / table", "Extracted on", "Method", "Confidence"]):
        hdr(ws, r + 1, 1 + j, t)
    lin = cfg.get("lineage", [])
    for k, d in enumerate(lin):
        for j, key in enumerate(["item", "source", "page_table", "extracted_on", "method", "confidence"]):
            rc(ws, r + 2 + k, 1 + j, d.get(key, ""), kind="input", border=True, wrap=True)
    widths(ws, {"A": 34, "B": 70, "C": 18, "D": 16, "E": 22, "F": 14})

    for nm in names:
        S_[nm].sheet_view.zoomScale = 90
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.save(path)
    return path


# ----------------------------------------------------------------------------
# Sample data (ILLUSTRATIVE ONLY - not real fund disclosures)
# ----------------------------------------------------------------------------
UNIVERSE = [  # name, sector, mcap, illustrative base weight (%)
    ("HDFC Bank", "Financials", "Large", 9.0), ("ICICI Bank", "Financials", "Large", 8.0),
    ("Axis Bank", "Financials", "Large", 3.5), ("State Bank of India", "Financials", "Large", 3.0),
    ("Kotak Mahindra Bank", "Financials", "Large", 2.5), ("Bajaj Finance", "Financials", "Large", 2.4),
    ("HDFC Life Insurance", "Financials", "Large", 0.8), ("SBI Life Insurance", "Financials", "Large", 0.8),
    ("Infosys", "IT", "Large", 5.0), ("TCS", "IT", "Large", 3.8), ("HCL Technologies", "IT", "Large", 1.8),
    ("Tech Mahindra", "IT", "Large", 1.0), ("Reliance Industries", "Energy", "Large", 7.0),
    ("Coal India", "Energy", "Large", 0.9), ("NTPC", "Power", "Large", 2.0), ("Power Grid Corp", "Power", "Large", 1.5),
    ("Bharti Airtel", "Telecom", "Large", 4.5), ("Larsen & Toubro", "Industrials", "Large", 4.0),
    ("Hindustan Unilever", "FMCG", "Large", 2.5), ("ITC", "FMCG", "Large", 3.5), ("Nestle India", "FMCG", "Large", 0.9),
    ("Asian Paints", "Materials", "Large", 1.1), ("UltraTech Cement", "Materials", "Large", 1.4),
    ("Tata Steel", "Metals", "Large", 1.0), ("Hindalco", "Metals", "Large", 1.0),
    ("Maruti Suzuki", "Automobile", "Large", 2.0), ("Mahindra & Mahindra", "Automobile", "Large", 3.0),
    ("Bajaj Auto", "Automobile", "Large", 1.0), ("Titan Company", "Consumer Discretionary", "Large", 1.8),
    ("Trent", "Consumer Discretionary", "Large", 1.0), ("Sun Pharma", "Healthcare", "Large", 2.0),
    ("Dr Reddy's Laboratories", "Healthcare", "Large", 0.9), ("Cipla", "Healthcare", "Large", 0.9),
    ("Persistent Systems", "IT", "Mid", 0), ("Max Healthcare", "Healthcare", "Mid", 0),
    ("Cummins India", "Industrials", "Mid", 0), ("Federal Bank", "Financials", "Mid", 0),
    ("Indian Hotels", "Consumer Discretionary", "Mid", 0), ("Voltas", "Consumer Discretionary", "Mid", 0),
]
MC = {u[0]: u[2] for u in UNIVERSE}
SEC = {u[0]: u[1] for u in UNIVERSE}


def _rows(w):
    return [{"stock": k, "sector": SEC[k], "mcap": MC[k], "weight": v} for k, v in w.items()]


def _gen_fund(rng, hold_p, sigma, mid_share, equity):
    raw = {}
    for name, sec, mc, b in UNIVERSE:
        if mc == "Large":
            if b >= 3.0 or rng.random() < hold_p:
                raw[name] = b * rng.lognormvariate(0, sigma)
        elif mid_share > 0 and rng.random() < 0.75:
            raw[name] = rng.uniform(1, 3)
    large = [k for k in raw if MC[k] == "Large"]
    mid = [k for k in raw if MC[k] != "Large"]
    if not mid:
        mid_share = 0
    sl, sm = sum(raw[k] for k in large), sum(raw[k] for k in mid)
    w = {k: raw[k] / sl * equity * (1 - mid_share) for k in large}
    w.update({k: raw[k] / sm * equity * mid_share for k in mid})
    for _ in range(20):  # cap each stock at 9.5% and redistribute the excess
        over = {k: v for k, v in w.items() if v > 0.095}
        if not over:
            break
        excess = sum(v - 0.095 for v in over.values())
        rest = {k: v for k, v in w.items() if v <= 0.095}
        tr = sum(rest.values())
        w = {**{k: 0.095 for k in over}, **{k: v + excess * v / tr for k, v in rest.items()}}
    w = {k: round(v, 4) for k, v in w.items() if v >= 0.002}
    return w, round(1 - sum(w.values()), 4)


def _gen_prev(rng, cur):
    raw = {k: v * rng.lognormvariate(0, 0.30) for k, v in cur.items() if not (rng.random() < 0.10 and v < 0.03)}
    cands = [u[0] for u in UNIVERSE if u[0] not in cur]
    for k in rng.sample(cands, min(2, len(cands))):
        raw[k] = rng.uniform(0.005, 0.02)
    scale = sum(cur.values()) / sum(raw.values())
    return {k: round(v * scale, 4) for k, v in raw.items()}


def sample_config():
    rng = random.Random(2026)
    profiles = [
        ("Fund A", "Sample Large Cap Fund A", 0.60, 0.65, 0.00, 0.975, 0.0110, 0.0140, 0.45, 400000, "Manager A"),
        ("Fund B", "Sample Bluechip Fund B", 0.97, 0.08, 0.00, 0.990, 0.0125, 0.0140, 0.20, 300000, "Manager B"),
        ("Fund C", "Sample Large & Midcap Style Fund C", 0.50, 0.75, 0.22, 0.960, 0.0170, 0.0140, 1.10, 200000, "Manager C"),
        ("Fund D", "Sample Focused Large Cap Fund D", 0.38, 0.85, 0.08, 0.975, 0.0100, 0.0140, 0.60, 100000, "Manager D"),
    ]
    funds = []
    for sh, nm, hp, sg, ms, eq, ex, cx, to, inv, mg in profiles:
        w, cash = _gen_fund(rng, hp, sg, ms, eq)
        pw = _gen_prev(rng, w)
        funds.append({"short": sh, "name": nm, "category": "Large Cap", "mandate_large_min": 0.80,
                      "aum_cr": round(rng.uniform(8000, 45000)), "expense": ex, "cat_expense": cx, "turnover": to,
                      "cash": cash, "investment": inv, "manager": mg, "holdings": _rows(w), "prev_holdings": _rows(pw)})
    bw = {u[0]: u[3] for u in UNIVERSE if u[2] == "Large"}
    tot = sum(bw.values())
    bw = {k: round(v / tot, 4) for k, v in bw.items()}
    top = max(bw, key=bw.get)
    bw[top] = round(bw[top] + (1 - sum(bw.values())), 4)
    return {
        "report_title": "FactSheetAI - Overlap & Drift Report",
        "as_of": "Sep-2026", "prior_month": "Mar-2026", "benchmark_name": "Illustrative Large-Cap Index",
        "source_note": "AMFI factsheets and AMC monthly portfolio disclosures (replace with your uploaded files)",
        "data_status": "SAMPLE / ILLUSTRATIVE DATA - fund names and weights are made up to demo the workbook. Replace with real disclosures.",
        "funds": funds, "benchmark": _rows(bw),
        "lineage": [{"item": f["name"], "source": "sample_input.json", "page_table": "n/a",
                     "extracted_on": "2026-10-19", "method": "Synthetic sample", "confidence": "n/a"} for f in funds],
    }


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--input")
    ap.add_argument("--output", default="FactSheetAI_Report.xlsx")
    ap.add_argument("--dump-sample")
    a = ap.parse_args()
    if a.dump_sample:
        with open(a.dump_sample, "w", encoding="utf-8") as sample_file:
            json.dump(sample_config(), sample_file, indent=1)
        print("ILLUSTRATIVE sample written to", a.dump_sample)
    if not a.input:
        ap.error("--input is required to build a report; sample_config() is illustrative only")
    with open(a.input, encoding="utf-8") as input_file:
        cfg = json.load(input_file)
    print("saved", build(cfg, a.output))
