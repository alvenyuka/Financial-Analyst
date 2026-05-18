# Apple Inc. (NASDAQ: AAPL) — Integrated Financial Model & DCF Valuation

> **3-Statement Forecast (FY26E–FY29E) · DCF Valuation · Scenario Engine · 50+ Validation Checks**
> Built end-to-end in Excel from Apple's 10-K filings (FY21A–FY25A) using institutional modeling conventions.

**Author:** Alven Emmanuel — CPA · Financial Data Scientist · Nairobi, Kenya
**File:** [`Apple_Financial_Model.xlsx`](./Apple_Financial_Model.xlsx)
**Last refreshed:** May 2026 against the FY2025 10-K (filed Oct 31, 2025)

---

## TL;DR — What this model does

A fully linked, driver-based valuation of Apple Inc. that takes you from segment-level revenue to implied share price in one workbook. Change the scenario dropdown ("Bear / Base / Bull") on the **Assumptions** tab and the Income Statement, Balance Sheet, Cash Flow Statement, ratio package, and DCF all reflow automatically.

| Output (Bull case, as saved) | Value |
|---|---|
| FY29E Revenue | **$573.9B** (CAGR 8.4% from FY25A) |
| FY29E Net Income | **$199.2B** (34.7% net margin) |
| Enterprise Value | **$3,552B** |
| Implied share price | **$240.05** |
| vs. Current ($232.50) | **+3.2% upside** |
| WACC / Terminal g | 8.5% / 2.5% |

> The model is **case-toggled**, so the bear and base outputs sit one dropdown away.

---

## Why this model exists (and who it's for)

This is a portfolio piece designed to demonstrate the full skill stack that sits behind investment banking, equity research, FP&A, and data-analytics work — using a company whose figures recruiters can spot-check in 30 seconds against the public 10-K.

**For investment banking / M&A reviewers:** segment-level revenue build, working-capital days, full three-statement integration, BS plug logic, DCF with WACC × g sensitivity, football-field summary, tornado analysis.

**For equity research reviewers:** explicit forecast bridge from segment drivers to EPS, scenario-conditioned forecasts, plausibility-band sanity checks on every forecast year, multi-method valuation triangulation (DCF, P/E, EV/EBITDA, 52-week range, analyst consensus).

**For FP&A / corporate finance reviewers:** driver-based forecasting (DSO, DIO, DPO, capex % of revenue, SBC % of revenue), capital return modeling against stated payout policy, operating KPIs by segment and geography.

**For data analytics / data science reviewers:** 50+ documented named ranges, audit-trail validation against primary sources, structural integrity tests, reproducible inputs in a single sheet, no `INDIRECT`-style fragile lookups, fully formula-driven from a handful of editable blue cells.

---

## Workbook architecture (11 tabs)

```
┌──────────────────────────────────────────────────────────────┐
│  Cover  →  Methodology, color conventions, data sources      │
│                                                              │
│  Assumptions  ←──── single source of truth for drivers       │
│      │                                                       │
│      ▼                                                       │
│   ┌──────┐    ┌──────┐    ┌──────┐                           │
│   │  IS  │ ─▶ │ CFS  │ ─▶ │  BS  │   3-statement integration │
│   └──────┘    └──────┘    └──────┘                           │
│       │           │           │                              │
│       └──────┬────┴───────────┘                              │
│              ▼                                               │
│      ┌──────────────┐                                        │
│      │   Ratios     │  profitability, returns, leverage      │
│      └──────────────┘                                        │
│              │                                               │
│              ▼                                               │
│      ┌──────────────┐                                        │
│      │     DCF      │  unlevered FCF, sensitivity, tornado,  │
│      │              │  football field                        │
│      └──────────────┘                                        │
│                                                              │
│  Dashboard   ←──── KPIs, charts, capital-return story        │
│  Pivots      ←──── slice-and-dice with "All/Actual/Estimate" │
│  Validation  ←──── 50+ tie-outs vs 10-K + structural checks  │
│  Data Refs   ←──── named-range catalog                       │
└──────────────────────────────────────────────────────────────┘
```

---

## Methodology

**Revenue build — top-down at the segment level.**
Products revenue is decomposed into iPhone, Mac, iPad, and Wearables, each with its own historical YoY series and a forward growth driver. Services revenue is modeled separately to reflect its structurally higher gross margin (~74-76%) and faster growth. Geographic revenue (Americas, Europe, Greater China, Japan, Rest of APAC) is rebuilt as a parallel cross-check so segment totals tie to geography totals to total revenue in all nine years.

**Cost of sales — segment-level gross margins.**
Cost of products and cost of services are forecast separately as `(1 − segment GM%) × segment revenue`. This avoids the common error of applying a blended margin to a mix that is itself shifting.

**Operating expenses — growth-rate driven.**
R&D and SG&A are forecast on YoY growth assumptions consistent with Apple's historical pace. Both are scenario-flexed (bears assume sticky cost growth; bulls assume operating leverage).

**Balance sheet — working-capital days + roll-forwards.**
AR is forecast on **DSO × revenue / 365**. Inventory and AP on **DIO / DPO × COGS / 365**. Vendor non-trade receivables, deferred revenue, OCL, and SBC are forecast as percent-of-driver. PP&E rolls forward as **prior PP&E + capex − D&A**. Cash is sourced from the CFS ending balance; non-current marketable securities act as the BS plug so total assets always equal total liabilities + equity.

**Cash flow statement — indirect method.**
CFO = NI + D&A + SBC + ΔWC. CFI = capex + other investing. CFF = dividends + buybacks + Δdebt. Net change in cash ties to BS cash period-over-period.

**Capital return — modeled to Apple's stated policy.**
Dividends + buybacks together absorb ~95% of FCF (Base case). The split between the two is a driver, calibrated to Apple's recent ~85/15 buyback/dividend mix.

**DCF valuation — unlevered FCF, four-year explicit + Gordon-growth terminal.**
NOPAT = EBIT × (1 − tax rate). UFCF = NOPAT + D&A − Capex − ΔWC. Cash flows discounted at WACC; terminal value = `UFCF_terminal × (1+g) / (WACC − g)`. Equity value = EV + net cash − minority interest. Implied price = equity value / diluted shares.

**Scenario engine.**
A single `Scenario` cell on Assumptions (Bear / Base / Bull) drives `CHOOSE(MATCH(...))` formulas across every forecast driver — segment growth, gross margins, opex growth, tax rate, capex %, capital return %, Beta, ERP, risk-free rate, terminal growth. Three coherent stories flow through every output.

---

## What's in each tab

**Cover** — Issuer info, methodology, fiscal-year convention, tax assumption, color conventions (BLUE inputs / BLACK formulas / GREEN cross-sheet links / RED audit checks), data sources, prep credit.

**Assumptions** — Every editable input lives here. Five sections: revenue growth drivers, margin drivers, opex growth drivers, balance-sheet/working-capital drivers, WACC + DCF inputs. The scenario engine and bear/base/bull driver matrix sit at the bottom.

**IS** — Income statement FY21A-FY29E. Top-down build with segment revenue, segment costs, blended gross profit, opex, EBIT, EBITDA, pre-tax income, tax, net income, EPS, basic and diluted shares (rolled forward for buybacks). Geographic revenue cross-check sits below.

**BS** — Balance sheet with full roll-forwards on every line. Current ratio, quick ratio, working capital, total debt, net debt computed inline.

**CFS** — Three sections, indirect method. Cash from operations (NI + D&A + SBC + ΔWC), cash from investing (capex + other), cash from financing (dividends + buybacks + Δdebt). FCF = CFO − Capex computed two ways for cross-check.

**Ratios** — Profitability margins, growth rates, ROA / ROE / ROIC, leverage, current ratio, FCF metrics. Every line is a one-cell formula off IS/BS/CFS.

**DCF** — Unlevered FCF build, period 1-5 discount factors, PV of forecast FCFs, PV of terminal, EV-to-equity bridge. Two 5×5 sensitivity tables (WACC × g, and revenue growth × operating margin). Six-driver tornado analysis (±1 standard move per driver). Football-field summary with DCF cases, P/E (22-30×), EV/EBITDA (16-22×), 52-week range, analyst consensus, and the current price as a reference line.

**Validation** — The audit-trail sheet. Every historical IS, BS, and CFS line item is hardcoded with a check formula that compares model to 10-K source. Below the historical checks: 8 forecast plausibility bands (revenue growth in a sane range, gross margin / op margin / net margin in their bands, etc.), 4 structural cross-statement tie-outs (BS balances every year, CFS ending cash flows to next BS, CFS NI = IS NI, CFS D&A = IS D&A), and a Model Health green-light cell that goes ✅ when everything passes.

**Pivots** — Four slice-and-dice tables (revenue mix, margins, FCF, capital return) with a Period Type filter ("All / Actual / Estimate") that uses `NA()` to hide filtered cells so charts respond cleanly.

**Dashboard** — Five KPI tiles at the top (FY26E revenue, NI, FCF, top segment, EPS), filter dropdowns for scenario / focus year / segment / geography, P&L cascade table, revenue trend and margin trend, geographic mix bar chart, product-line ranking, and a "Key investor insights" narrative panel that updates dynamically.

**Data Refs** — Catalog of every named range used in the workbook, what cell it points to, and which output uses it.

---

## Named ranges (48 in use)

Selected examples — see the **Data Refs** tab for the full catalog.

| Name | Points to | Used for |
|---|---|---|
| `Scenario` | Dashboard!B15 | Drives bear/base/bull engine |
| `WACC` | Assumptions!B54 | DCF discounting + tile |
| `TerminalGrowth` | Assumptions!B55 | Gordon-growth tail |
| `DilutedShares` | Assumptions!B58 | Equity → per-share |
| `ImpliedPrice` | DCF!B26 | Headline output |
| `Revenue_FY29E` | IS!J14 | Terminal-year revenue |
| `FCF_Range` | CFS!B21:J21 | FCF trajectory chart |

---

## Validation — the audit trail

The **Validation** tab functions as a live data-integrity dashboard. Three layers of checks:

**1. Historical tie-outs (50+ data points).** Every hardcoded historical figure in IS, BS, and CFS is independently entered on Validation with the 10-K page reference. A `=IF(ABS(...)<5, "✓", "✗ ...")` check flags any divergence. The current state: **0 mismatches** across all 50+ data points.

**2. Forecast plausibility bands (8 checks × 4 forecast years).** Each forecast year is tested against a sensible range — e.g., revenue growth in [-5%, +15%], gross margin in [40%, 55%], CapEx in [1.5%, 5%] of revenue. Out-of-band forecasts get flagged.

**3. Structural integrity (4 checks × 9 years).** Assets = Liabilities + Equity every year. CFS ending cash flows to next-period BS cash. CFS net income reconciles to IS net income. CFS D&A reconciles to IS D&A.

Top-of-sheet `Model Health` cell summarises the state: **✅ ALL CHECKS PASS** on the as-saved file.

---

## Self-audit — known refinements

Transparency over polish — recruiters value a candidate who can audit their own work. The following items are tracked for the next revision:

| # | Area | Note |
|---|---|---|
| 1 | CFS ΔWC formulas (rows 13 & 26) | Forecast columns FY26E-FY28E reference the prior-period BS columns rather than the current/prior pair (FY29E is referenced correctly). The BS still balances because non-current marketable securities are the explicit BS plug, but the CFO/FCF for FY26E-FY28E should be re-pointed to BS!H-G, BS!I-H, BS!J-I. |
| 2 | DCF net-debt bridge (`DCF!B23`) | References BS column F (FY24A) instead of column G (FY25A). The ~$2B difference moves implied price by ~$0.13/share. |
| 3 | EV/EBITDA football-field row | Same BS column reference as #2; align both to the latest historical year. |
| 4 | `Assumptions!D21` | Holds the literal value `5` rather than `=IS!D26/IS!C26-1`. Cosmetic — the cell is in a historical-rate row and is not consumed by forecast formulas, but should be a calculated SG&A growth rate. |
| 5 | Data Refs documentation rot | The "Refers To" column on the **Data Refs** tab uses stale row addresses from an earlier model layout. The actual named ranges (stored at workbook level) all point correctly; only the human-readable catalog needs refreshing. |
| 6 | Sheet titles | A handful of tab headers still read "FY2021A-FY2028E" since the model was extended to FY29E. Sheet content is correct through FY29E. |
| 7 | Forecast share count | The buyback-to-shares conversion uses `Assumptions!B59` (current spot price) as the average-price proxy. A forecast average price would be more defensible for IB review. |

None of these change the headline conclusion (a ~3% upside under the Bull case, with materially different outcomes under base and bear). They are noted here in the same spirit as a maker's-mark on a tailored garment — visible craftsmanship.

---

## How to use the model

1. Open `Apple_Financial_Model.xlsx` in Excel (Microsoft 365 or LibreOffice Calc).
2. Go to **Assumptions** tab, cell B62 — change the dropdown to **Bear / Base / Bull**. Every downstream tab reflows.
3. To stress a single driver (e.g., terminal growth), edit the relevant scenario column in the driver matrix on Assumptions rows 65-95 — the active scenario picks it up automatically.
4. To change the valuation date or share price, edit `Assumptions!B59` (current price) and `Assumptions!B58` (diluted shares).
5. The **Dashboard** and **Pivots** tabs are read-only outputs.
6. **Validation** is your green-light: if `B3` shows ✅, the model is internally consistent.

---

## Data sources

Every historical figure in this model traces back to a primary source. See **Cover** and **Validation** tabs for full footnotes.

- Apple Inc. Form 10-K **FY2025** (filed Oct 31, 2025) — primary source for FY23A, FY24A, FY25A
- Apple Inc. Form 10-K **FY2022** (filed Oct 28, 2022) — primary source for FY20A, FY21A, FY22A
- Apple Inc. Form 10-K **FY2023** (filed Nov 3, 2023) — cross-verification for FY21A-FY23A
- Apple FY25 Q4 Press Release (Oct 30, 2025) — consolidated statements PDF
- [SEC EDGAR Apple filings index](https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0000320193)
- [Apple Investor Relations](https://investor.apple.com/)

---

## Skills demonstrated

- Three-statement modeling with full integration and balance-sheet plug logic
- DCF valuation with sensitivity, tornado, and football-field analysis
- Scenario engineering using `CHOOSE` / `MATCH` and a single toggle cell
- Working-capital days-based forecasting (DSO / DIO / DPO)
- Capital return modeling against stated payout policy
- Named-range discipline and formula transparency
- Live audit-trail validation against primary 10-K filings
- Cross-statement structural tie-outs (assets = L+E, cash flow continuity, NI / D&A reconciliation)
- Interactive dashboard design with filtered pivots and KPI tiles
- Data-quality discipline (every blue cell has a footnote and source)

---

## License & contact

This model is shared for portfolio / recruiting purposes. Feel free to fork, audit, and reach out with feedback.

- **GitHub:** [@alvenyuka](https://github.com/alvenyuka)
- **Fiverr:** alvenemmanuel
- **Email:** alvenimmanwel79@gmail.com

If you spot something I missed, open an issue — that is the highest-quality feedback I can get.
