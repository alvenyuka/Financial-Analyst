# Apple Inc. (AAPL) — Three-Statement Model & DCF Valuation

> Three-statement model (IS, BS, CFS) and DCF valuation built from Apple's SEC 10-K filings and quarterly press releases. Every historical line reconciles back to its source filing; the balance sheet balances to zero in all 9 years (5 historical + 4 projected).

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](../LICENSE)
[![Excel](https://img.shields.io/badge/Excel-LibreOffice%20compatible-217346?logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![No macros](https://img.shields.io/badge/macros-none-success)](#modelling-conventions)
[![Validation](https://img.shields.io/badge/validation-50%2F50%20data%20points-success)](#validation)

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
3. Review the **Assumptions** tab (colour-coded inputs) — rows 37–47 hold the DCF drivers (WACC, terminal growth, tax rate).
4. Confirm the **Validation** tab is fully green before trusting any output.
5. Cross-check any historical line against the source filing in `Source_filings/`.

## Workbook tabs

| Tab | Contents |
|---|---|
| Dashboard | Headline metrics |
| Cover | Model scope and version |
| Assumptions | Single source of truth for every driver — segment growth, margins, WACC, terminal growth, tax rate |
| IS / BS / CFS | Income statement, balance sheet, cash flow statement — 5 years historical (FY21–FY25) + 4 years projected (FY26E–FY29E) |
| Ratios | Derived ratios |
| DCF | Unlevered FCF build, WACC discounting, Gordon Growth terminal value, two 5×5 sensitivity tables |
| Validation | Line-by-line 10-K cross-reference, structural integrity checks, forecast sanity bands |
| Pivots / Data Refs | Supporting pivot tables and reference data |

## Validation

The Validation tab cross-references every hardcoded historical figure against Apple's SEC 10-K filings and press releases — **50 of 50 data points match (100%)**, zero mismatches. Verified May 2026 against 10-Ks filed Oct 2022, Nov 2023, and Oct 2025, cross-checked against SEC EDGAR (CIK 0000320193) and Apple's investor-relations press releases.

- **Balance sheet balances to zero** (Assets − Liabilities − Equity = 0) in all 9 years, historical and projected.
- **Cash flow ties to the balance sheet**: ending cash in each period equals the next period's beginning cash.
- **Cross-statement consistency**: net income and D&A match between the CFS and IS in every year.
- **Forecast sanity bands** (FY26E–FY29E): revenue growth, gross margin, operating margin, tax rate, net margin, CapEx %, FCF margin, and liquidity all fall inside pre-defined plausibility ranges — e.g. revenue growth 8.2–8.5% (band: −5% to +15%), operating margin 36.5–39.8% (band: 25–40%).

Bundled 10-K line items (e.g. "Other current liabilities") are footnoted in the Validation tab with the components that sum to the model figure, so the audit trail survives Apple's own line-item aggregation choices.

## DCF summary

Base case (WACC 8.5%, terminal growth 2.5%): **implied share price ≈ $240** vs. a reference price of $232.50 at model build time — roughly 3% upside. Two sensitivity tables are wired up on the DCF tab: implied price by WACC × terminal growth, and by revenue growth × operating margin, so the output moves visibly as assumptions change rather than being presented as a single point estimate.

## Modelling Conventions

### One source of truth for inputs

The Assumptions tab holds every driver. Every other tab recalculates from it. Inputs are colour-coded; calculated cells are locked.

### Every historical figure traces to its source

The model ships alongside the 10-Ks and press releases the figures were taken from, in `Source_filings/`. The Validation tab reconciles each historical line back to the filing it came from, down to the page number.

### Structural integrity is enforced

The Validation tab runs the balance-sheet identity, the cash-flow ties, and inter-statement consistency checks before the model is treated as correct — see the [Validation](#validation) section above for the actual results, not just the methodology.

### Excel-native, no add-ins

No macros, no external data feeds, no add-ins. Opens in Excel or LibreOffice Calc.

## License

MIT — see [`LICENSE`](../LICENSE).

## Connect

📫 [alvenyuka2@gmail.com](mailto:alvenyuka2@gmail.com) · 💼 [LinkedIn](https://www.linkedin.com/in/alven-yuka-610b78174/) · 🐙 [GitHub](https://github.com/alvenyuka)
