# Apple Inc. (AAPL): Three-Statement Model & DCF Valuation

> Three-statement model (IS, BS, CFS) and DCF valuation built from Apple's SEC 10-K filings and quarterly press releases. Every historical line carries a filing and page citation; the balance sheet balances to zero in all 9 years (5 historical + 4 projected). The DCF tab has known defects, listed under Known issues below.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](../LICENSE)
[![Excel](https://img.shields.io/badge/Excel-LibreOffice%20compatible-217346?logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![No macros](https://img.shields.io/badge/macros-none-success)](#modelling-conventions)
[![Validation](https://img.shields.io/badge/validation-50%2F50%20lines%20reconciled-success)](#validation)

## Project Structure

```
Apple/
├── Apple_Financial_Model.xlsx     # Three-statement model, DCF, scenario toggles
├── Source_filings/                # 10-K, 10-Q, press releases backing the model
└── README.md
```

## Quick Start

1. Open `Apple_Financial_Model.xlsx` in Excel or LibreOffice Calc.
2. Start on the **Dashboard** tab for the headline numbers.
3. Review the **Assumptions** tab (colour-coded inputs). The DCF drivers sit lower down than the tab's own notes suggest: risk-free rate, ERP and beta are rows 44 to 47, the tax rate used for WACC is row 50, WACC itself is `B54`, terminal growth is `B55`, the share count is `B58`, the reference price is `B59`, and the scenario toggle is `B62`. Rows 37 to 47 hold CapEx, D&A and capital-return drivers. `DCF!A40` and `DCF!A42` both point at the wrong rows; see Known issues below.
4. Confirm the **Validation** tab is fully green before trusting any output.
5. Cross-check any historical line against the source filing in `Source_filings/`.

## Workbook tabs

| Tab | Contents |
|---|---|
| Dashboard | Headline metrics |
| Cover | Model scope and version |
| Assumptions | Single source of truth for every driver: segment growth, margins, WACC, terminal growth, tax rate |
| IS / BS / CFS | Income statement, balance sheet, cash flow statement: 5 years historical (FY21–FY25) + 4 years projected (FY26E–FY29E) |
| Ratios | Derived ratios |
| DCF | Unlevered FCF build, WACC discounting, Gordon Growth terminal value, two 5×5 sensitivity tables |
| Validation | Line-by-line 10-K cross-reference, structural integrity checks, forecast sanity bands |
| Pivots / Data Refs | Supporting pivot tables and reference data |

## Validation

The Validation tab reconciles **50 of 50 line items (100%)** between its own transcription of the figures and the `IS`, `BS` and `CFS` tabs. Fifty counts line items, each covering five fiscal years, so 250 cells.

Be precise about what that reconciliation is. Column G compares the Validation tab's copy of a figure against the statement tab's copy of the same figure. No cell in the workbook, and no line in `validate_model.py`, opens a PDF. If both transcriptions came from the same misreading of a filing, all 50 would still show a tick. The tie to the filings is carried by the Source column's page citations and by a manual check made in May 2026 against the 10-Ks filed Oct 2022, Nov 2023 and Oct 2025, cross-checked against SEC EDGAR (CIK 0000320193) and Apple's investor-relations press releases. That check is real; it is just not the thing the tick marks measure.

- **Balance sheet balances to zero** (Assets − Liabilities − Equity = 0) in all 9 years, historical and projected.
- **Cash flow ties to the balance sheet**: ending cash in each period equals the next period's beginning cash.
- **Cross-statement consistency**: net income and D&A match between the CFS and IS in every year.
- **Forecast sanity bands** (FY26E–FY29E): revenue growth, gross margin, operating margin, tax rate, net margin, CapEx %, FCF margin, and liquidity all fall inside pre-defined plausibility ranges, e.g. revenue growth 8.2–8.5% (band: −5% to +15%), operating margin 36.5–39.8% (band: 25–40%).

Bundled 10-K line items (e.g. "Other current liabilities") are footnoted in the Validation tab with the components that sum to the model figure, so the audit trail survives Apple's own line-item aggregation choices.

## DCF summary

The workbook is saved on the **Bull** scenario (`Assumptions!B62`). On that scenario it computes **WACC 7.79%** (`B54`, cost of equity and after-tax cost of debt at 90/10 weights) and **terminal growth 3.0%** (`B55`), and `DCF!B26` returns an **implied share price of $240.05** against the $232.50 reference in `B59`, about 3% upside.

Earlier versions of this file, the repo README and the profile README all described that as a base case of WACC 8.5% with terminal growth 2.5%. It is not. Those are not the inputs the model runs, and they are not any scenario in its scenario engine either. The workbook's own caption at `DCF!A40` says 8.5% / 2.5%, which is where the error came from, and the workbook itself is not edited here. **At WACC 8.5% and g 2.5% the model's own sensitivity grid returns $205.92** (`DCF!D36`), which against $232.50 is roughly 11% downside rather than 3% upside. The direction of the call depends on which pair a reader believes, so both are stated.

Two sensitivity tables are wired up on the DCF tab: implied price by WACC x terminal growth, and by revenue growth x operating margin, so the output moves visibly as assumptions change rather than being presented as a single point estimate.

### Known issues on the DCF tab and elsewhere in the workbook

These are defects in the workbook, listed rather than quietly corrected: fixing them changes published valuation figures, and that is a decision to take deliberately.

- **The terminal value is discounted one period too many.** The explicit forecast is four years (`B15:E15` = 1, 2, 3, 4) and the Gordon Growth terminal value in `F13` is a value as at the end of year 4, so it belongs discounted by `(1+WACC)^4`. `F15` is 5 and `F16` discounts by `(1+WACC)^5`. Terminal value is 85% of enterprise value, $3.02tn of $3.55tn, so this is not a rounding matter. The sensitivity grid twenty rows below uses `^4` on the same inputs, so the tab computes one valuation two ways and gets two answers.
- **The footnotes point at the wrong cells.** `DCF!A40` says to change inputs on "Assumptions sheet rows 37-47" and `DCF!A42` says the grid assumes `WACC = Assumptions!B46, terminal g = Assumptions!B47, tax = Assumptions!B42`. Every one is off by about eight rows; the formulas in that block reference `Assumptions!$B$54` and `$B$55`.
- **The football field's DCF rows are a different model.** `DCF - Base case` at $144.91 / $154.14 / $163.66 reads from the second sensitivity grid, a revenue-growth-by-operating-margin build, not from the main DCF. The tab therefore publishes three irreconcilable DCF answers: $240.05 headline, $205.92 at the caption's stated inputs, $154.14 in the football field.
- **The P/E and EV/EBITDA bars are undiscounted.** They apply multiples to FY29E earnings and EBITDA and compare the result to today's $232.50 without discounting back four years, which is most of why those bars sit at $262 to $398, far above every DCF bar.
- **The reference price has no date.** `Assumptions!B59` holds $232.50 with nothing recording when that was the market price, so the comparison cannot be checked as of any particular day.
- **`Assumptions!D21` is a hardcoded 5.** Row 21 is SG&A expense YoY growth, and `C21`, `E21` and `F21` are formulas off the income statement. `D21` (FY2023) is the literal number 5, displaying as 500% growth; the figure from the model's own IS is -0.65%. It does not reach the valuation, because the forecast columns use the scenario `CHOOSE`, which is exactly why neither the in-sheet validation nor `validate_model.py` catches it.
- **`DCF!A23` is labelled "- Net debt" and the figure is net cash.** `B23` is cash plus investments minus total debt, which is +50,021, and `B24` adds it. The arithmetic is right for a company in a net cash position; the label is not, and equity value exceeding enterprise value by exactly that amount reads as a sign error until a reader opens the formula.
- **Two Dashboard chart blocks labelled FY25A read FY24A.** The Region block at `B97:C102` and the Product block at `B108:C113` read column F of the Dashboard's own tables, where column F is FY24A and column G is FY25A. The segment block between them reads `IS!F8` and `IS!F13`, where column F really is FY25A. So the Dashboard shows Services at 109.2 in one chart and 96.17 in another, both labelled FY25A, three rows apart.
- **Three different tax rates, none reconciled in writing.** `Assumptions!B25` is a 13.0% effective rate for the forecast, `Assumptions!B50` is 15% for WACC, and `DCF!B8` computes NOPAT at `(1 - 0.85)`, also 15%. A normalised marginal rate for NOPAT alongside an effective rate on the income statement is standard practice, but it needs saying, and the 13.0% forecast rate is below every one of the five historical years (13.3, 16.2, 14.7, 24.1, 15.6) with no stated reason.
- **The forecast margin step-change is the DCF's main value driver and is justified nowhere.** Products gross margin is an input that jumps from FY25A 36.8% to a flat 40.5% across FY26E to FY29E, and services from 75.4% to 78.0%, taking operating margin from 32.0% to 36.5% rising to 39.8% against five years of history in a 29.8% to 32.0% band. The sanity band quoted above for operating margin is 25% to 40%, which a forecast reaching 39.8% cannot fail, so passing it does not mean much.
- **`IS!A1` reads "FY2021A to FY2028E"** while the columns run to FY2029E.

## Modelling Conventions

### One source of truth for inputs

The Assumptions tab holds every driver. Every other tab recalculates from it. Inputs are colour-coded; calculated cells are locked.

### Every historical figure traces to its source

The filings behind FY23 onward ship alongside the model in `Source_filings/`; FY21 and FY22 cite the FY2022 and FY2023 10-Ks, linked to SEC EDGAR on the Validation tab. The Validation tab's Source column names the filing and page for each historical line, which is the audit trail. The tick marks in column G are a separate thing and check tab against tab, as set out under Validation above.

`Source_filings/` also carries `FY26_Q1_Consolidated_Financial_Statements.pdf`, which nothing in the model cites: FY26 is a projection here, and the Validation tab's sources are the FY2022, FY2023 and FY2025 10-Ks plus two Q4 press releases. Checking one quarter of FY26 actuals against the FY26E forecast would be a good use of it.

### Structural integrity is enforced

The Validation tab runs the balance-sheet identity, the cash-flow ties, and inter-statement consistency checks before the model is treated as correct. See the [Validation](#validation) section above for the actual results, not just the methodology.

### Excel-native, no add-ins

No macros, no external data feeds, no add-ins. Opens in Excel or LibreOffice Calc.

## License

MIT. See [`LICENSE`](../LICENSE).

## Connect

📫 [alvenyuka2@gmail.com](mailto:alvenyuka2@gmail.com) · 💼 [LinkedIn](https://www.linkedin.com/in/alven-yuka-610b78174/) · 🐙 [GitHub](https://github.com/alvenyuka)
