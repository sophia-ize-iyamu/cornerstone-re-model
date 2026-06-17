"""Excel workbook and investment-memo writers for Cornerstone.

The workbook is written with the engine's computed values (the Python model in cornerstone.py is the
single source of truth), formatted as a clean Assumptions / Pro Forma / Returns / Sensitivity workbook.
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

NAVY, LIGHT = "1F3864", "D9E1F2"
hdr = PatternFill("solid", fgColor=NAVY); sub = PatternFill("solid", fgColor=LIGHT)
bw = Font(bold=True, color="FFFFFF"); thin = Side(style="thin", color="BFBFBF")
bd = Border(left=thin, right=thin, top=thin, bottom=thin)

AROWS = [
    ("Units", "units", "#,##0"), ("Avg monthly rent / unit", "rent_mo", "$#,##0"),
    ("Other income (yr 1)", "other_income", "$#,##0"), ("Vacancy + credit loss", "vacancy", "0.0%"),
    ("OpEx ratio (% EGI)", "opex_ratio", "0%"), ("NOI / income growth", "noi_growth", "0.0%"),
    ("Purchase price", "purchase_price", "$#,##0"), ("Acquisition cost %", "acq_cost_pct", "0.0%"),
    ("Capex reserve / unit", "capex_per_unit", "$#,##0"), ("Loan-to-value", "ltv", "0%"),
    ("Mortgage rate", "rate", "0.00%"), ("Amortization (yrs)", "amort_years", "#,##0"),
    ("Hold (yrs)", "hold_years", "#,##0"), ("Exit cap rate", "exit_cap", "0.00%"),
    ("Disposition cost %", "sale_cost_pct", "0.0%"),
]

def _hrow(ws, r, vals, money_from=1):
    for j, v in enumerate(vals, 1):
        c = ws.cell(r, j, v); c.font = bw; c.fill = hdr; c.border = bd
        c.alignment = Alignment(wrap_text=True, vertical="center")

def write_excel(a, R, sens, here):
    p = R["p"]; rows = p["rows"]
    wb = Workbook(); ws = wb.active; ws.title = "Assumptions"
    ws["A1"] = "Cornerstone | Assumptions"; ws["A1"].font = Font(bold=True, size=13, color=NAVY)
    ws.column_dimensions["A"].width = 30; ws.column_dimensions["B"].width = 16
    for i, (lab, k, fmt) in enumerate(AROWS, start=2):
        ws.cell(i, 1, lab); c = ws.cell(i, 2, a[k]); c.number_format = fmt; c.fill = sub; c.border = bd
    der = [("Loan", p["loan"]), ("Acquisition costs", p["acq_costs"]), ("Equity", p["equity"]),
           ("Capex reserve / yr", p["capex"]), ("Debt service (annual)", p["pmt"])]
    for i, (lab, v) in enumerate(der, start=len(AROWS) + 3):
        ws.cell(i, 1, lab); c = ws.cell(i, 2, round(v)); c.number_format = "$#,##0"; c.font = Font(bold=True)

    m = wb.create_sheet("Pro Forma")
    H = ["Year", "GPR", "EGI", "OpEx", "NOI", "Debt svc", "Interest", "Principal", "Balance",
         "DSCR", "CFADS", "Levered CF", "Unlevered CF"]
    _hrow(m, 1, H)
    m.column_dimensions["A"].width = 6
    for col in "BCDEFGHIKLM": m.column_dimensions[col].width = 12
    m.column_dimensions["J"].width = 7
    # Year 0
    m.cell(2, 1, 0)
    m.cell(2, 9, round(p["loan"]))                 # balance = loan
    m.cell(2, 12, round(R["lev"][0]))              # -equity
    m.cell(2, 13, round(R["unlev"][0]))            # -(price+acq)
    for r in rows:
        t = r["t"]; row = t + 2
        vals = [t, r["gpr"], r["egi"], r["opex"], r["noi"], r["pmt"], r["interest"], r["principal"],
                r["balance"], r["dscr"], r["cfads"], R["lev"][t], R["unlev"][t]]
        for j, v in enumerate(vals, 1):
            c = m.cell(row, j, round(v) if j != 10 and j != 1 else (round(v, 2) if j == 10 else v))
            c.number_format = "0.00" if j == 10 else ("#,##0" if j > 1 else "0")
    last = len(rows) + 2
    rb = last + 2
    block = [("Levered IRR", R["lev_irr"], "0.0%"), ("Unlevered IRR", R["unlev_irr"], "0.0%"),
             ("Equity multiple", R["equity_multiple"], '0.00"x"'), ("Min DSCR", R["min_dscr"], '0.00"x"'),
             ("Going-in cap rate", R["going_in_cap"], "0.00%"), ("Debt yield", R["debt_yield"], "0.00%"),
             ("Avg cash-on-cash", R["avg_coc"], "0.0%"), ("Exit value", p["exit_value"], "$#,##0")]
    m.cell(rb, 1, "RETURNS").font = Font(bold=True, color=NAVY)
    for i, (lab, v, fmt) in enumerate(block, start=rb + 1):
        m.cell(i, 1, lab); c = m.cell(i, 2, round(v) if fmt.startswith("$") else v)
        c.number_format = fmt; c.font = Font(bold=True)

    s = wb.create_sheet("Sensitivity")
    caps, grid = sens
    s["A1"] = "Levered IRR sensitivity"; s["A1"].font = Font(bold=True, size=12, color=NAVY)
    s["A2"] = "Rows: rent-growth shift (ppt)   |   Columns: exit-cap shift (bp)"
    s.cell(4, 1, "growth \\ exit cap").font = Font(bold=True)
    for j, dc in enumerate(caps, start=2):
        c = s.cell(4, j, f"{int(dc*100):+d}bp"); c.font = bw; c.fill = hdr; c.alignment = Alignment(horizontal="center")
    for i, (dr, line) in enumerate(grid):
        rr = 5 + i
        lc = s.cell(rr, 1, f"{dr:+.0f} ppt"); lc.font = bw; lc.fill = hdr
        for j, val in enumerate(line, start=2):
            c = s.cell(rr, j, val if val is not None else 0); c.number_format = "0.0%"; c.border = bd
            if dr == 0 and caps[j - 2] == 0: c.fill = sub; c.font = Font(bold=True)
    s.column_dimensions["A"].width = 16
    wb.save(os.path.join(here, "cornerstone_model.xlsx"))

def write_memo(a, R, sens, here):
    os.makedirs(os.path.join(here, "out"), exist_ok=True)
    p = R["p"]; rows = p["rows"]; caps, grid = sens
    g = lambda i, j: (f"{grid[i][1][j]:.1%}" if grid[i][1][j] is not None else "n/a")
    md = f"""# Investment memo: Project Cornerstone

**Asset:** {a['units']}-unit stabilized multifamily property
**Ask:** approve ${p['equity']/1e6:.1f}M equity for a ${a['purchase_price']/1e6:.0f}M acquisition ({a['ltv']:.0%} LTV)

## Recommendation
Proceed. At a ${a['purchase_price']/1e6:.0f}M basis (a **{R['going_in_cap']:.2%} going-in cap**, above the
{a['rate']:.2%} debt rate, so leverage is accretive), the deal underwrites to a **{R['lev_irr']:.1%} levered IRR**
and a **{R['equity_multiple']:.2f}x equity multiple** over a {a['hold_years']}-year hold, with a minimum
**DSCR of {R['min_dscr']:.2f}x** and a **{R['debt_yield']:.2%} debt yield**. The exit is held flat to entry at a
{a['exit_cap']:.2%} cap, so the return is driven by NOI growth and amortization, not cap-rate compression.

## Returns
| Metric | Value |
|---|---|
| Going-in cap rate | {R['going_in_cap']:.2%} |
| Year-1 NOI | ${rows[0]['noi']/1e6:.2f}M |
| Unlevered IRR | {R['unlev_irr']:.1%} |
| Levered IRR | {R['lev_irr']:.1%} |
| Equity multiple | {R['equity_multiple']:.2f}x |
| Avg cash-on-cash | {R['avg_coc']:.1%} |
| Min DSCR | {R['min_dscr']:.2f}x |
| Debt yield | {R['debt_yield']:.2%} |
| Exit value (@ {a['exit_cap']:.2%}) | ${p['exit_value']/1e6:.1f}M |

## Pro forma (NOI build)
| Year | NOI ($mm) | Debt svc ($mm) | DSCR | Levered CF ($mm) | Cash-on-cash |
|---|---|---|---|---|---|
""" + "\n".join(
        f"| {r['t']} | {r['noi']/1e6:.2f} | {r['pmt']/1e6:.2f} | {r['dscr']:.2f} | {r['lev_cf']/1e6:.2f} | {r['coc']:.1%} |"
        for r in rows) + f"""

## Levered IRR sensitivity (rent growth vs exit cap)
| growth \\ exit cap | {int(caps[0]*100):+d}bp | {int(caps[1]*100):+d}bp | {int(caps[2]*100):+d}bp | {int(caps[3]*100):+d}bp | {int(caps[4]*100):+d}bp |
|---|---|---|---|---|---|
| {grid[0][0]:+.0f} ppt | {g(0,0)} | {g(0,1)} | {g(0,2)} | {g(0,3)} | {g(0,4)} |
| {grid[1][0]:+.0f} ppt | {g(1,0)} | {g(1,1)} | **{g(1,2)}** | {g(1,3)} | {g(1,4)} |
| {grid[2][0]:+.0f} ppt | {g(2,0)} | {g(2,1)} | {g(2,2)} | {g(2,3)} | {g(2,4)} |

## Key risks
- **Exit cap expansion** is the dominant risk; the sensitivity above isolates it against rent growth.
- **Rent growth** below {a['noi_growth']:.1%} compresses returns; underwriting holds it conservative.
- **Rate / refinancing:** debt is sized to a {R['min_dscr']:.2f}x DSCR and {R['debt_yield']:.2%} debt yield for cushion.

## Method and references
Standard real-estate underwriting: pro forma to NOI, cap-rate valuation, amortizing debt with DSCR and
debt yield, and a forward-NOI exit. Following Geltner & Miller, *Commercial Real Estate Analysis and
Investments*; NCREIF and ULI market conventions; cap-rate / NOI / DSCR as used in Argus Enterprise.

*Figures illustrative; built on transparent assumptions (see the Assumptions tab). The method, not the
specific levels, is the point.*
"""
    open(os.path.join(here, "out", "investment_memo.md"), "w", encoding="utf-8").write(md)
