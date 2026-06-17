# Cornerstone: a real-estate acquisition underwriting model

What does a real-estate investment banking or acquisitions analyst build to underwrite a deal? A pro
forma from potential rent down to NOI. A cap-rate valuation. A financing structure with DSCR and debt
yield. A hold-period cash flow. An exit on a forward cap rate. And the returns an investment committee
asks for: unlevered and levered IRR, equity multiple, and cash-on-cash. Cornerstone does all of it from
first principles in **Excel and Python**, and writes a one-page investment memo.

Modelled asset: a **200-unit stabilized multifamily property**, acquired for $47M.

## What it does
- **Builds the pro forma** from gross potential rent through vacancy, other income, and operating
  expenses to **NOI**, grown over the hold.
- **Values on a cap rate** (going-in and exit), and computes the **going-in cap** and **debt yield**.
- **Sizes and amortizes the debt**, with an annual **DSCR** so coverage is visible.
- **Runs the hold-period cash flow** and an **exit on forward NOI / exit cap**, net of disposition costs.
- **Computes the returns**: unlevered IRR, levered IRR, equity multiple, and average cash-on-cash.
- **Runs a sensitivity** of levered IRR against the two factors that move a real-estate return most:
  **exit cap rate and rent growth**.
- Writes `cornerstone_model.xlsx` (Assumptions / Pro Forma / Returns / Sensitivity) and
  `out/investment_memo.md`.

## Headline (base case)
$47M purchase at a **6.0% going-in cap** (above the 5.25% debt rate, so leverage is accretive):
**12.6% levered IRR**, **8.6% unlevered**, a **1.74x equity multiple** over 5 years, **5.5% average
cash-on-cash**, and a **1.50x minimum DSCR** on a $19.7M equity cheque. The exit is held flat to entry
at a 6.0% cap, so the return comes from NOI growth and amortization, not cap-rate compression.

## The data
Real deals come off the rent roll and the broker offering memorandum; here the figures are
**illustrative, representative of a stabilized multifamily asset**. The point is the method: the pro
forma, the cap-rate math, the amortizing debt, the exit, and the return build. Change an assumption and
the model and returns recompute.

## Run it
Requires Python and `openpyxl`.
```bash
pip install -r requirements.txt
python cornerstone.py
```
You will see the underwrite in the console, and the script regenerates `cornerstone_model.xlsx` and
`out/investment_memo.md`.

## Concepts and references
- **NOI, cap rate, going-in vs exit:** the core of real-estate valuation; value = NOI / cap rate.
- **DSCR and debt yield:** lender coverage and leverage tests on the loan.
- **Levered vs unlevered IRR, equity multiple, cash-on-cash:** the standard real-estate return suite.
- Methodology follows Geltner & Miller, *Commercial Real Estate Analysis and Investments*; NCREIF and
  ULI market conventions; and the cap-rate / NOI / DSCR framework used in Argus Enterprise.

## Limitations
The asset is synthetic and the assumptions are illustrative, so the figures are not a real deal. Rent is
grown at a single rate rather than a unit-level rent roll, taxes and depreciation are not modelled (it is
a pre-tax underwrite), debt is a simple amortizing loan (no IO period or refinancing), and the exit is a
single forward-cap assumption. The point is to show I can underwrite a real-estate acquisition end to
end, pull the returns an IC needs, and frame the decision, in Excel and in code.
