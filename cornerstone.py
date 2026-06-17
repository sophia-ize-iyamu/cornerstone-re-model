"""
Cornerstone: a real-estate acquisition underwriting model.

What a real-estate investment banking / acquisitions analyst builds to underwrite a deal: a pro forma
(potential rent down to NOI), a cap-rate valuation, a financing structure with DSCR and debt yield,
a hold-period cash flow, an exit on a forward cap rate, and the returns an investment committee asks
for, unlevered and levered IRR, equity multiple, and cash-on-cash. Cornerstone does all of it from
first principles, writes a formula-driven Excel workbook, and prints an investment summary.

Modelled asset: a 200-unit stabilized multifamily property. The figures are illustrative, built on
transparent, defensible assumptions; real deals come off the rent roll and broker OM. What is real is
the method: the pro forma, the cap-rate math, the amortizing debt, the exit, and the return build.

Run:  python cornerstone.py   (requires openpyxl)
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

HERE = os.path.dirname(os.path.abspath(__file__))

A = {
    "units":          200,
    "rent_mo":        2000.0,    # average in-place monthly rent per unit
    "other_income":   150_000.0, # parking, fees, etc. (year 1)
    "vacancy":        0.05,      # vacancy + credit loss
    "opex_ratio":     0.40,      # operating expenses as % of EGI
    "noi_growth":     0.035,     # annual growth of rents / income
    "purchase_price": 47_000_000.0,
    "acq_cost_pct":   0.02,      # closing / acquisition costs
    "capex_per_unit": 300.0,     # annual capital reserve per unit
    "ltv":            0.60,      # loan-to-value on purchase price
    "rate":           0.0525,    # mortgage rate
    "amort_years":    30,        # amortization
    "hold_years":     5,
    "exit_cap":       0.060,     # exit (going-out) cap rate, flat to entry (conservative)
    "sale_cost_pct":  0.02,      # disposition costs
}

def proforma(a):
    loan = a["purchase_price"] * a["ltv"]
    acq_costs = a["purchase_price"] * a["acq_cost_pct"]
    equity = a["purchase_price"] + acq_costs - loan
    capex = a["capex_per_unit"] * a["units"]
    r, n = a["rate"], a["amort_years"]
    pmt = loan * r / (1 - (1 + r) ** -n)        # annual amortizing payment

    rows, bal = [], loan
    for t in range(1, a["hold_years"] + 1):
        g = (1 + a["noi_growth"]) ** (t - 1)
        gpr = a["units"] * a["rent_mo"] * 12 * g
        egi = gpr * (1 - a["vacancy"]) + a["other_income"] * g
        opex = a["opex_ratio"] * egi
        noi = egi - opex
        interest = bal * r
        principal = pmt - interest
        bal -= principal
        cfads = noi - capex                      # cash flow after reserves, before debt
        lev_cf = cfads - pmt                      # levered cash flow to equity
        rows.append(dict(t=t, gpr=gpr, egi=egi, opex=opex, noi=noi, capex=capex,
                         interest=interest, principal=principal, balance=bal, pmt=pmt,
                         dscr=noi / pmt, cfads=cfads, lev_cf=lev_cf,
                         coc=lev_cf / equity))
    # Exit on forward NOI / exit cap
    exit_noi = rows[-1]["noi"] * (1 + a["noi_growth"])
    exit_value = exit_noi / a["exit_cap"]
    net_sale = exit_value * (1 - a["sale_cost_pct"])
    net_equity_proceeds = net_sale - rows[-1]["balance"]
    return dict(rows=rows, loan=loan, equity=equity, acq_costs=acq_costs, capex=capex, pmt=pmt,
                exit_noi=exit_noi, exit_value=exit_value, net_sale=net_sale,
                net_equity_proceeds=net_equity_proceeds)

def irr(flows, lo=-0.95, hi=2.0):
    def npv(x): return sum(cf / (1 + x) ** i for i, cf in enumerate(flows))
    a, b = npv(lo), npv(hi)
    if a * b > 0: return None
    for _ in range(200):
        m = (lo + hi) / 2; fm = npv(m)
        if abs(fm) < 1: return m
        if a * fm < 0: hi, b = m, fm
        else: lo, a = m, fm
    return (lo + hi) / 2

def npv(rate, flows): return sum(cf / (1 + rate) ** i for i, cf in enumerate(flows))

def returns(a):
    p = proforma(a)
    rows, eq = p["rows"], p["equity"]
    # unlevered: -(price+acq) at t0; NOI-capex each year; + net sale at exit
    unlev = [-(a["purchase_price"] + p["acq_costs"])]
    for i, rr in enumerate(rows):
        cf = rr["cfads"]
        if i == len(rows) - 1: cf += p["net_sale"]
        unlev.append(cf)
    # levered: -equity at t0; levered cf each year; + net equity proceeds at exit
    lev = [-eq]
    for i, rr in enumerate(rows):
        cf = rr["lev_cf"]
        if i == len(rows) - 1: cf += p["net_equity_proceeds"]
        lev.append(cf)
    em = sum(lev[1:]) / eq
    return dict(p=p, unlev=unlev, lev=lev,
                unlev_irr=irr(unlev), lev_irr=irr(lev), equity_multiple=em,
                avg_coc=sum(rr["coc"] for rr in rows) / len(rows),
                going_in_cap=rows[0]["noi"] / a["purchase_price"],
                debt_yield=rows[0]["noi"] / p["loan"],
                min_dscr=min(rr["dscr"] for rr in rows))

def sensitivity(base):
    caps = [-0.50, -0.25, 0.0, 0.25, 0.50]   # exit cap shift, % (50bp steps)
    rents = [-1.0, 0.0, 1.0]                  # rent-growth shift, ppt
    grid = []
    for dr in rents:
        line = []
        for dc in caps:
            a = dict(base); a["exit_cap"] = base["exit_cap"] + dc / 100
            a["noi_growth"] = base["noi_growth"] + dr / 100
            line.append(returns(a)["lev_irr"])
        grid.append((dr, line))
    return caps, grid

if __name__ == "__main__":
    R = returns(A)
    print("== Cornerstone: multifamily acquisition underwrite ==")
    print(f"Price ${A['purchase_price']/1e6:.1f}M | Loan ${R['p']['loan']/1e6:.1f}M ({A['ltv']:.0%} LTV) | Equity ${R['p']['equity']/1e6:.1f}M")
    print(f"Year-1 NOI ${R['p']['rows'][0]['noi']/1e6:.2f}M | Going-in cap {R['going_in_cap']:.2%} | Debt yield {R['debt_yield']:.2%}")
    print(f"Exit value ${R['p']['exit_value']/1e6:.1f}M @ {A['exit_cap']:.2%} | Min DSCR {R['min_dscr']:.2f}x")
    print(f"Unlevered IRR {R['unlev_irr']:.1%} | Levered IRR {R['lev_irr']:.1%} | Equity multiple {R['equity_multiple']:.2f}x | Avg cash-on-cash {R['avg_coc']:.1%}")
    import cornerstone_io
    cornerstone_io.write_excel(A, R, sensitivity(A), HERE)
    cornerstone_io.write_memo(A, R, sensitivity(A), HERE)
    print("\nWrote cornerstone_model.xlsx and out/investment_memo.md")
