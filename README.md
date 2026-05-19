![Financial Analyst Portfolio](banner.svg)

<p align="center">
  <img alt="Excel" src="https://img.shields.io/badge/Excel-217346?logo=microsoftexcel&logoColor=white">
  <img alt="DCF" src="https://img.shields.io/badge/DCF-valuation-2E5C8A">
  <img alt="Three-Statement" src="https://img.shields.io/badge/Three--Statement-model-1F4E79">
  <img alt="Scenario" src="https://img.shields.io/badge/Scenario-analysis-5B7553">
  <img alt="License" src="https://img.shields.io/badge/license-MIT-blue">
</p>

---

## Overview

Company financial analyses built from primary-source filings: three-statement models, DCF valuations, scenario analysis. Each workbook ships with the filings it was sourced from, a validation tab that reconciles every historical line to its source, and a short write-up of the figures that matter.

I'm Alven, a CPA finalist (Kenya) awaiting ICPAK membership. This repository is where I publish modelling work as I build out the technical side of accounting and finance.

## Current models

| Company | Ticker | Status | Folder |
|---|---|---|---|
| Apple Inc. | AAPL | Complete | `Apple/` |

More to come.

## Modelling conventions

A few conventions run through every analysis in this repository.

**One source of truth for inputs.** A dedicated tab holds assumptions and inputs. Every other tab recalculates from it. Inputs are colour-coded; calculated cells are locked.

**Every historical figure traces to its source.** Each workbook ships alongside the 10-K, 10-Q, and press releases the figures were taken from. The validation tab reconciles each historical line back to the filing it came from.

**Structural integrity is enforced.** The validation tab runs the balance-sheet identity (assets = liabilities + equity), the cash flow ties (net income to retained earnings, depreciation to PPE), and inter-statement consistency checks. The workbook does not treat a model as correct until those checks pass.

**Scenario toggles where they help.** Bear, base, and bull scenarios are wired up on questions where the answer changes materially across them. Where it does not, the model stays single-case to keep the audit trail clean.

Files open in Excel or LibreOffice Calc. No macros, no external data feeds, no add-ins.

## Folder structure

```
Financial-Analyst/
├── README.md
├── banner.svg
└── Apple/                 Three-statement model, DCF, source filings, notes
```

Additional company analyses will follow this same structure.

## Contact

GitHub: [@alvenyuka](https://github.com/alvenyuka)  
Email: alvenyuka2@gmail.com  
Fiverr: alvenemmanuel

If you spot an error in any model, open an issue. If you'd like work in this style on a company you care about, send a note.

## License

MIT.
