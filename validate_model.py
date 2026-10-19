"""Independently re-derive every accounting identity in a model workbook.

The workbook already has a Validation tab that prints "ALL CHECKS PASS". This
script exists because that tick mark is produced by the same spreadsheet whose
correctness is in question. A formula that is wrong reports a tick just as
confidently as one that is right, so a reader has no way to distinguish the two
without recomputing the arithmetic outside the file.

This does that. It adds the raw line items up in Python and compares the result
against the totals the workbook reports. Nothing here reads a cell whose value
is a checkmark.

Scope, stated up front because an earlier version of this file overstated it.
The line items live on the Validation tab, which is a second transcription of
the statements rather than the statements themselves. Re-deriving the identities
there proves the transcription is internally consistent and proves nothing about
`IS`, `BS` or `CFS`. That gap was real: with every numeric cell on the IS and BS
replaced by the number 1, this script still printed "19 of 19 ... matched" and
exited 0. The transcription checks added below close it, by reading the source
sheets directly and comparing them to the Validation tab, row by row and year by
year. Both families must pass for the run to succeed.

The valuation is checked too, without reusing the workbook's own cell choices:
the change in working capital is rebuilt from the balance-sheet lines and compared
with the cash flow statement; the forecast balance sheet is re-added from its
components; and the DCF is rebuilt from the valuation date (Assumptions!B60) and
the fiscal year-end dates, so a cash flow from a year that ended before the
valuation date is excluded and its cash is counted in net cash instead. WACC,
free cash flow, present values, terminal value, enterprise and equity value, the
per-share price, the sensitivity grid and the discounted multiples are all
re-derived. Still out of scope: the `Ratios`, `Dashboard` and `Pivots` tabs, and
the Bear and Bull prices, which are recorded with each scenario active and printed
as recorded rather than re-derived.

What it checks:

* **Income statement**, down the page. Revenue from its two segments, cost of
  sales from its two segments, then gross profit, operating income, pre-tax
  income and net income, each from the line above it.
* **Balance sheet**, both sides, ending with the identity everything rests on:
  assets minus liabilities minus equity is zero.
* **Cash flow statement**, including the two ties that are easy to assert and
  easy to get wrong: net change in cash equals the three sections added up, and
  ending cash equals beginning cash plus that change.
* **Cross-statement ties**, which is where a three-statement model usually breaks.
  The cash flow statement's ending cash has to equal the balance sheet's cash for
  the same year, and it has to equal the next year's beginning cash. A model can
  satisfy every within-statement check and still fail these.
* **Transcription**, one check per Validation row, five years each: the figure on
  the Validation tab against the cell or cells it was copied from on `IS`, `BS`
  or `CFS`. The source cells are read out of the Validation tab's own column G
  formulas, so the mapping comes from the workbook and the comparison is done
  here. Without this, every identity above could pass on a workbook whose
  statements had been emptied.

Usage:

    python validate_model.py                        # defaults to the Apple model
    python validate_model.py path/to/Model.xlsx

Exits 0 when every check passes, 1 otherwise, so it can run in CI.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is required: pip install openpyxl")

DEFAULT_MODEL = Path(__file__).parent / "Apple" / "Apple_Financial_Model.xlsx"

# Money is in millions of USD. The workbook rounds to whole millions, so an exact
# equality test would fail on rounding alone. One million on figures in the
# hundreds of billions is about four parts in a million.
TOLERANCE = 1.0

# Cells on the Assumptions tab that the ratio and DCF rebuilds read. Several are
# formulas driven by the scenario switch in B62, so the Base-case value is only
# given as an example.
RISK_FREE_RATE = "B45"           # risk-free rate (Base case 4.5%)
EQUITY_RISK_PREMIUM = "B46"      # equity risk premium (Base case 5.0%)
BETA = "B47"                     # Apple's beta (Base case 1.1)
PRE_TAX_COST_OF_DEBT = "B49"     # pre-tax cost of debt (Base case 4.0%)
WACC_TAX_RATE = "B50"            # tax rate used in the WACC (Base case 15%; a formula, =G25)
TARGET_DEBT_TO_CAPITAL = "B52"   # debt share of the capital structure; equity is the rest (Base case 10%)
WACC_CELL = "B54"                # the WACC the workbook calculated, which the rebuild is checked against
TERMINAL_GROWTH = "B55"          # terminal growth rate g (Base case 2.5%)
DILUTED_SHARES = "B58"           # diluted shares outstanding, in millions
REFERENCE_PRICE = "B59"          # share price on the valuation date
VALUATION_DATE = "B60"           # date of that price; cash flows are discounted back to it
ACTIVE_SCENARIO = "B62"          # Base, Bear or Bull: the scenario switch
PE_MID_MULTIPLE = "C97"          # middle P/E multiple of the football field
EV_EBITDA_MID_MULTIPLE = "C98"   # middle EV/EBITDA multiple of the football field


def _wrap(text: str, width: int):
    """Split text into lines of at most `width` characters, breaking between words."""
    words, line, out = text.split(), "", []
    for w in words:
        if len(line) + len(w) + 1 > width:
            out.append(line)
            line = w
        else:
            line = f"{line} {w}".strip()
    if line:
        out.append(line)
    return out


class Check:
    """One identity, evaluated for every year in the model."""

    def __init__(self, label: str, expected, actual, years, note: str | None = None,
                 advisory: bool = False, tol: float = TOLERANCE):
        self.label = label
        self.tol = tol
        self.expected = expected
        self.actual = actual
        self.years = years
        #: Context a reader needs to interpret a failure, printed alongside it.
        #: A check that fails without explanation invites the reader to assume
        #: either "broken model" or "ignore it", and one of those is wrong.
        self.note = note
        #: An advisory check is reported but does not fail the run. It compares
        #: two figures whose equality depends on a presentation choice rather
        #: than on arithmetic, so a difference is a question, not a defect.
        self.advisory = advisory

    @property
    def failures(self):
        """Years where the expected and actual values differ by more than the tolerance, or one is missing."""
        out = []
        for year, exp, act in zip(self.years, self.expected, self.actual):
            if exp is None or act is None:
                out.append((year, exp, act, "missing value"))
            elif abs(exp - act) > self.tol:
                gap = f"{exp - act:,.1f}" if self.tol >= 1 else f"{exp - act:+.6f}"
                out.append((year, exp, act, f"differs by {gap}"))
        return out

    @property
    def passed(self) -> bool:
        """True when no year fails."""
        return not self.failures


def read_rows(ws, first_col=2, n_years=5):
    """Map each row label in column A to its list of yearly values.

    Raises if a label carrying numbers appears more than once. Silently keeping
    the first occurrence would make every check on that label report PASS for a
    row nobody chose, which is worse than not checking it: the result looks
    verified.
    """
    rows = {}
    where = {}
    duplicates = []

    for r in range(1, ws.max_row + 1):
        label = ws.cell(r, 1).value
        if not isinstance(label, str):
            continue
        label = label.strip()
        values = [ws.cell(r, c).value for c in range(first_col, first_col + n_years)]
        if not any(isinstance(v, (int, float)) for v in values):
            continue
        if label in rows:
            duplicates.append((label, where[label], r))
            continue
        rows[label] = [v if isinstance(v, (int, float)) else None for v in values]
        where[label] = r

    if duplicates:
        detail = "\n".join(
            f"    {label!r}: rows {first} and {second}" for label, first, second in duplicates
        )
        raise ValueError(
            "Duplicate numeric row labels in the Validation tab. Every check on "
            "these would silently validate the first one and report a pass:\n"
            + detail
        )
    return rows


def find_section_headers(ws):
    """Locate each statement's year-header row from its section title, so inserted
    rows on the Validation tab cannot shift a check onto the wrong row."""
    titles = {"income statement": "INCOME STATEMENT", "balance sheet": "BALANCE SHEET",
              "cash flow statement": "CASH FLOW STATEMENT"}
    found = {}
    for r in range(1, ws.max_row + 1):
        text = str(ws.cell(r, 1).value or "")
        for key, title in titles.items():
            if key not in found and text.startswith(title):
                found[key] = r + 1
    missing = [k for k in titles if k not in found]
    if missing:
        raise ValueError(f"section title(s) not found on the Validation tab: {missing}")
    return found


def check_year_headers(ws, header_rows, n_years=5):
    """Every statement must be laid out over the same years.

    The years come from the income statement's header. If another statement used
    a different set, the comparisons would still line up positionally and mean
    nothing.
    """
    headers = {}
    for name, r in header_rows.items():
        years_in_header = []
        for c in range(2, 2 + n_years):
            years_in_header.append(ws.cell(r, c).value)
        headers[name] = tuple(years_in_header)
    distinct = set(headers.values())
    if len(distinct) > 1:
        detail = "\n".join(f"    {name}: {list(cols)}" for name, cols in headers.items())
        raise ValueError(
            "The statements are not laid out over the same years, so a "
            "column-by-column comparison would be meaningless:\n" + detail
        )
    return list(distinct.pop())


def get_row(rows, label):
    """Return one Validation row's yearly values, or fail loudly if the label is gone."""
    if label not in rows:
        raise KeyError(f"row not found in the Validation tab: {label!r}")
    return rows[label]


# The next three work on lists of yearly values. A year where any input is
# missing (None) gives None, so the check reports "missing value" for that year
# instead of crashing on None + number.

def add_lists(*lists):
    """Add the lists together, year by year."""
    totals = []
    for i in range(len(lists[0])):
        values = [one_list[i] for one_list in lists]
        if None in values:
            totals.append(None)
        else:
            totals.append(sum(values))
    return totals


def subtract_lists(first, *others):
    """first minus the sum of the other lists, year by year."""
    results = []
    for i in range(len(first)):
        start = first[i]
        to_take_off = [one_list[i] for one_list in others]
        if start is None or None in to_take_off:
            results.append(None)
        else:
            results.append(start - sum(to_take_off))
    return results


def divide_lists(top, bottom):
    """top divided by bottom, year by year."""
    results = []
    for i in range(len(top)):
        if top[i] is None or bottom[i] is None:
            results.append(None)
        else:
            results.append(top[i] / bottom[i])
    return results


def add_rows(rows, *labels):
    """Add the Validation rows with these labels, year by year."""
    return add_lists(*[get_row(rows, label) for label in labels])


def subtract_rows(rows, first_label, *other_labels):
    """The first row minus the other rows, year by year."""
    others = [get_row(rows, label) for label in other_labels]
    return subtract_lists(get_row(rows, first_label), *others)


def income_statement_checks(rows, years):
    """Each income statement line rebuilt from the line above it."""
    return [
        Check("IS  revenue = products + services",
              add_rows(rows, "Products revenue", "Services revenue"),
              get_row(rows, "Total revenue (calc)"), years),
        Check("IS  cost of sales = products + services",
              add_rows(rows, "Cost of products", "Cost of services"),
              get_row(rows, "Total cost of sales"), years),
        Check("IS  gross profit = revenue - cost of sales",
              subtract_rows(rows, "Total revenue (calc)", "Total cost of sales"),
              get_row(rows, "Gross profit (calc)"), years),
        Check("IS  opex = R&D + SG&A",
              add_rows(rows, "R&D expense", "SG&A expense"),
              get_row(rows, "Total operating expenses"), years),
        Check("IS  operating income = gross profit - opex",
              subtract_rows(rows, "Gross profit (calc)", "Total operating expenses"),
              get_row(rows, "Operating income (calc)"), years),
        Check("IS  pre-tax = operating income + other",
              add_rows(rows, "Operating income (calc)", "Other income/(expense)"),
              get_row(rows, "Pre-tax income (calc)"), years),
        Check("IS  net income = pre-tax - tax",
              subtract_rows(rows, "Pre-tax income (calc)", "Income tax expense"),
              get_row(rows, "Net income (calc)"), years),
    ]


def balance_sheet_checks(rows, years):
    """Each balance sheet subtotal re-added, ending with assets = liabilities + equity."""
    return [
        Check("BS  current assets sum to the subtotal",
              add_rows(rows, "Cash & equivalents", "Marketable securities (current)",
                       "Accounts receivable, net", "Inventory",
                       "Vendor non-trade receivables", "Other current assets"),
              get_row(rows, "Total current assets"), years),
        Check("BS  total assets = current + PP&E + other non-current",
              add_rows(rows, "Total current assets", "Marketable securities (non-curr)",
                       "Property, plant & equipment, net", "Other non-current assets"),
              get_row(rows, "Total assets"), years),
        Check("BS  current liabilities sum to the subtotal",
              add_rows(rows, "Accounts payable", "Other current liabilities", "Deferred revenue",
                       "Commercial paper", "Term debt (current)"),
              get_row(rows, "Total current liabilities"), years),
        Check("BS  total liabilities = current + LT debt + other non-current",
              add_rows(rows, "Total current liabilities", "Long-term debt (non-current)",
                       "Other non-current liabilities"),
              get_row(rows, "Total liabilities"), years),
        Check("BS  equity = common stock + retained earnings + other comprehensive income",
              add_rows(rows, "Common stock and APIC", "Retained earnings (deficit)",
                       "Accumulated OCI (loss)"),
              get_row(rows, "Shareholders' equity"), years),
        Check("BS  liabilities + equity subtotal",
              add_rows(rows, "Total liabilities", "Shareholders' equity"),
              get_row(rows, "Total liabilities + equity"), years),
        Check("BS  THE identity: assets = liabilities + equity",
              get_row(rows, "Total assets"), get_row(rows, "Total liabilities + equity"), years),
    ]


def cash_flow_checks(rows, years):
    """Each cash flow section re-added, and the sections added into the net change in cash."""
    return [
        Check("CF  investing = capex + other investing",
              add_rows(rows, "Capital expenditures", "Other investing"),
              get_row(rows, "Cash from investing"), years),
        Check("CF  financing = dividends + buybacks + other",
              add_rows(rows, "Dividends paid", "Share repurchases", "Other financing"),
              get_row(rows, "Cash from financing"), years),
        Check("CF  net change = operations + investing + financing",
              add_rows(rows, "Cash from operations", "Cash from investing", "Cash from financing"),
              get_row(rows, "Net change in cash"), years),
        Check("CF  ending cash = beginning cash + net change",
              add_rows(rows, "Beginning cash", "Net change in cash"),
              get_row(rows, "Ending cash"), years),
        Check("CF  free cash flow = operations - capex",
              add_rows(rows, "Cash from operations", "Capital expenditures"),
              get_row(rows, "Free cash flow (CFO - CapEx)"), years),
    ]


def cross_statement_checks(rows, years):
    """The ties between statements, which is where a three-statement model usually breaks."""
    ending_cash = get_row(rows, "Ending cash")
    beginning_cash = get_row(rows, "Beginning cash")

    cash_tie = Check(
        "TIE ending cash on the CFS = cash on the BS",
        ending_cash, get_row(rows, "Cash & equivalents"), years,
        note=(
            "Expected for FY21 to FY23: those cash flow statements reconcile "
            "to cash, cash equivalents and restricted cash, while the balance "
            "sheet line excludes restricted cash, so the gaps (989, 1,331 and "
            "772) are the restricted cash. From FY24 Apple's statements agree."
        ),
        advisory=True)

    # This year's beginning cash is last year's ending cash. A break here means
    # the statements are not actually linked, only arranged to look like they are.
    # There is one fewer pair than there are years, so the labels are built by hand.
    year_pairs = []
    for i in range(len(years) - 1):
        year_pairs.append(f"{years[i]} to {years[i + 1]}")
    roll_forward = Check("TIE beginning cash = prior year ending cash",
                         ending_cash[:-1], beginning_cash[1:], year_pairs)

    return [cash_tie, roll_forward]


def build_checks(rows, years):
    """The three-statement identities, each recomputed from its components in the Validation tab.

    `rows` maps each Validation row label to its yearly values. Every check adds or
    subtracts component rows and compares the result with the total the model
    reports, for every year.
    """
    checks = []
    checks += income_statement_checks(rows, years)
    checks += balance_sheet_checks(rows, years)
    checks += cash_flow_checks(rows, years)
    checks += cross_statement_checks(rows, years)
    return checks


# The Validation tab's column G holds, for each line item, a formula comparing
# that row against the sheet it was transcribed from. Two shapes are used:
#
#   =IF(SUMPRODUCT(ABS(B16:F16-IS!B8:F8))<1, ...)                 range form
#   =IF((ABS(B36-(BS!C5))+ABS(C36-(BS!D5))+...)<5, ...)           per-cell form
#
# Only the cell references are taken from them. The comparison itself is redone
# here in Python, so a broken formula in column G cannot hide a broken figure.
_RANGE_LINK = re.compile(
    r"ABS\((?P<vcol>[A-Z]+)(?P<vrow>\d+):[A-Z]+\d+"
    r"\s*-\s*(?P<sheet>[A-Za-z][A-Za-z0-9_ ]*)!(?P<scol>[A-Z]+)(?P<srow>\d+):[A-Z]+\d+\)"
)
_CELL_LINK = re.compile(
    r"ABS\((?P<vref>[A-Z]+\d+)\s*-\s*\(?(?P<src>[A-Za-z][A-Za-z0-9_ ]*![A-Z]+\d+"
    r"(?:\s*\+\s*[A-Za-z][A-Za-z0-9_ ]*![A-Z]+\d+)*)\)?\)"
)
_SHEET_CELL = re.compile(r"(?P<sheet>[A-Za-z][A-Za-z0-9_ ]*)!(?P<col>[A-Z]+)(?P<row>\d+)")


def _col_index(letters: str) -> int:
    """Column number of a column letter reference (A = 1, Z = 26, AA = 27)."""
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _formula_text(value) -> str | None:
    """Column G holds plain strings for some rows and ArrayFormula objects for
    others. Both carry the formula; only the attribute differs."""
    if isinstance(value, str):
        return value if value.startswith("=") else None
    text = getattr(value, "text", None)
    return text if isinstance(text, str) and text.startswith("=") else None


def read_source_links(ws_formulas, n_years=5):
    """Map each Validation row to the source cells it was transcribed from.

    Returns {row: (label, [(validation_cell, [(sheet, row, col), ...]), ...])},
    one entry per year column. Rows whose column G compares a source sheet to
    itself, such as the BALANCE CHECK row, carry no Validation-tab operand and
    are skipped: there is no transcription to check there.
    """
    links = {}
    for r in range(1, ws_formulas.max_row + 1):
        formula = _formula_text(ws_formulas.cell(r, 7).value)
        if not formula:
            continue
        label = ws_formulas.cell(r, 1).value
        if not isinstance(label, str):
            continue

        per_year = []
        match = _RANGE_LINK.search(formula)
        if match and int(match.group("vrow")) == r:
            vcol, scol = _col_index(match.group("vcol")), _col_index(match.group("scol"))
            sheet, srow = match.group("sheet"), int(match.group("srow"))
            for offset in range(n_years):
                per_year.append(((r, vcol + offset), [(sheet, srow, scol + offset)]))
        else:
            seen = set()
            for cell_match in _CELL_LINK.finditer(formula):
                vref = cell_match.group("vref")
                if vref in seen:
                    continue  # the TEXT() fallback repeats every term
                seen.add(vref)
                vsplit = re.match(r"([A-Z]+)(\d+)", vref)
                if int(vsplit.group(2)) != r:
                    continue
                sources = [
                    (m.group("sheet"), int(m.group("row")), _col_index(m.group("col")))
                    for m in _SHEET_CELL.finditer(cell_match.group("src"))
                ]
                per_year.append(((r, _col_index(vsplit.group(1))), sources))

        if len(per_year) == n_years:
            links[r] = (label.strip(), per_year)
    return links


def build_transcription_checks(wb_values, links, years):
    """One check per Validation row: the transcribed figure against its source.

    This is the part that actually opens IS, BS and CFS. Everything else in this
    script works off the Validation tab's own copy of the numbers.
    """
    validation = wb_values["Validation"]
    checks = []

    for _, (label, per_year) in sorted(links.items()):
        transcribed, derived = [], []
        for year_cell, sources in per_year:
            vrow, vcol = year_cell
            value = validation.cell(vrow, vcol).value
            transcribed.append(value if isinstance(value, (int, float)) else None)

            total = 0.0
            for sheet, srow, scol in sources:
                if sheet not in wb_values.sheetnames:
                    total = None
                    break
                cell = wb_values[sheet].cell(srow, scol).value
                if not isinstance(cell, (int, float)):
                    total = None
                    break
                total += cell
            derived.append(total)

        sheet_names = set()
        for _, sources in per_year:
            for source in sources:
                sheet_names.add(source[0])
        sheets = sorted(sheet_names)
        checks.append(
            Check(f"SRC {label} = {'/'.join(sheets)}", derived, transcribed, years,
                  note=(
                      "The Validation tab is a transcription of the statements, not "
                      "the statements. A difference here means the figure being "
                      "checked by every identity above is not the figure on the "
                      "statement it came from."
                  ))
        )
    return checks


RATIO_TOL = 1e-5  # 0.001 percentage points: covers the $1M rounding slack, catches any real error


def growth_rates(values):
    """Year-on-year growth. The first year has nothing before it, so it is None."""
    rates = [None]
    for i in range(1, len(values)):
        last_year, this_year = values[i - 1], values[i]
        if last_year is None or this_year is None:
            rates.append(None)
        else:
            rates.append(this_year / last_year - 1)
    return rates


def build_ratio_checks(wb_values, rows, years):
    """The Ratios tab, rebuilt from the Validation-tab line items rather than read from its formulas.

    The line items are the ones already traced to the filings (the SRC checks), so a
    ratio that passes here rests on the 10-K figures, not on whichever statement cell
    the Ratios formula happens to point at. Growth rates start in the second year.
    """
    ratios_sheet = wb_values["Ratios"]
    tax_rate = wb_values["Assumptions"][WACC_TAX_RATE].value
    reported = read_rows(ratios_sheet)

    # The line items, pulled out once under short names.
    revenue = get_row(rows, "Total revenue (calc)")
    gross_profit = get_row(rows, "Gross profit (calc)")
    operating_income = get_row(rows, "Operating income (calc)")
    net_income = get_row(rows, "Net income (calc)")
    depreciation = get_row(rows, "Depreciation & amortization")
    total_assets = get_row(rows, "Total assets")
    equity = get_row(rows, "Shareholders' equity")
    current_assets = get_row(rows, "Total current assets")
    current_liabilities = get_row(rows, "Total current liabilities")
    cash = get_row(rows, "Cash & equivalents")
    securities_current = get_row(rows, "Marketable securities (current)")
    securities_non_current = get_row(rows, "Marketable securities (non-curr)")
    operating_cash_flow = get_row(rows, "Cash from operations")
    free_cash_flow = get_row(rows, "Free cash flow (CFO - CapEx)")
    capex = get_row(rows, "Capital expenditures")

    # Total debt, and the combinations that several ratios share.
    debt = add_rows(rows, "Commercial paper", "Term debt (current)", "Long-term debt (non-current)")
    debt_plus_equity = add_lists(debt, equity)
    ebitda = add_lists(operating_income, depreciation)
    net_debt = subtract_lists(debt, cash)
    net_debt = subtract_lists(net_debt, securities_current)
    net_debt = subtract_lists(net_debt, securities_non_current)

    after_tax_operating_income = []
    for value in operating_income:
        after_tax_operating_income.append(None if value is None else value * (1 - tax_rate))

    # CapEx is stored as a negative cash outflow; the ratio is quoted as a positive share.
    capex_as_positive = []
    for value in capex:
        capex_as_positive.append(None if value is None else -value)

    derived = {
        "Gross margin %": divide_lists(gross_profit, revenue),
        "Operating margin %": divide_lists(operating_income, revenue),
        "Net margin %": divide_lists(net_income, revenue),
        "EBITDA margin %": divide_lists(ebitda, revenue),
        "Revenue growth %": growth_rates(revenue),
        "Operating income growth %": growth_rates(operating_income),
        "Net income growth %": growth_rates(net_income),
        "Return on assets (ROA)": divide_lists(net_income, total_assets),
        "Return on equity (ROE)": divide_lists(net_income, equity),
        "Return on invested capital": divide_lists(after_tax_operating_income, debt_plus_equity),
        "Debt / Equity": divide_lists(debt, equity),
        "Debt / Total capital": divide_lists(debt, debt_plus_equity),
        "Current ratio": divide_lists(current_assets, current_liabilities),
        "Net debt ($M)": net_debt,
        "Operating cash flow ($M)": operating_cash_flow,
        "Free cash flow ($M)": free_cash_flow,
        "FCF margin %": divide_lists(free_cash_flow, revenue),
        "FCF / Net income": divide_lists(free_cash_flow, net_income),
        "CapEx / Revenue": divide_lists(capex_as_positive, revenue),
    }

    checks = []
    for label, values in derived.items():
        if label not in reported:
            raise KeyError(f"row not found in the Ratios tab: {label!r}")
        start = 1 if values[0] is None else 0  # growth rates have no first year
        tol = TOLERANCE if label.endswith("($M)") else RATIO_TOL
        checks.append(Check(f"RAT {label}", values[start:], reported[label][start:],
                            years[start:], tol=tol))
    return checks


def _dcf_price(fcf, periods, wacc, g, net_cash, shares):
    """Per-share value: cash flows from years ending after the valuation date, plus a
    Gordon terminal value at the last forecast year-end."""
    pv = 0
    for cash_flow, t in zip(fcf, periods):
        if t > 0:  # a year that ended before the valuation date is already in net cash
            pv += cash_flow / (1 + wacc) ** t
    tv = fcf[-1] * (1 + g) / (wacc - g) / (1 + wacc) ** periods[-1]
    return (pv + tv + net_cash) / shares


# Labels the DCF rebuild depends on. An inserted row would otherwise move a cell
# reference silently, so each address is checked against its label first.
EXPECTED_LABELS = {
    ("IS", "A28"): "Operating income",
    ("CFS", "A10"): "Depreciation",
    ("CFS", "A13"): "Changes in working capital",
    ("CFS", "A19"): "Capital expenditures",
    ("CFS", "A30"): "Cash & equivalents, ending",
    ("BS", "B5"): "Cash",
    ("BS", "B15"): "Long-term marketable securities",
    ("BS", "B29"): "Long-term debt",
    ("Assumptions", "A25"): "Effective tax rate",
    ("Assumptions", "A54"): "WACC",
    ("Assumptions", "A60"): "Valuation date",
    ("DCF", "A14"): "Fiscal year end",
}

WC_ASSETS = (7, 8, 9, 10, 16)     # receivables, inventory, vendor receivables, other current, other non-current
WC_LIABS = (21, 22, 23, 30)       # payables, other current, deferred revenue, other non-current
BS_ASSETS = (5, 6, 7, 8, 9, 10, 14, 15, 16)
BS_LIABS_EQUITY = (21, 22, 23, 24, 25, 29, 30, 35, 36, 37)

# Forecast columns. IS, CFS and Assumptions hold FY2026E to FY2029E in G to J.
# The balance sheet starts one column later (H to K), so the prior year's balance
# sheet column is G to J. The DCF tab holds the same four years in B to E.
FORECAST_YEARS = ["FY26E", "FY27E", "FY28E", "FY29E"]
STATEMENT_COLUMNS = ["G", "H", "I", "J"]
BS_COLUMNS = ["H", "I", "J", "K"]
BS_PRIOR_COLUMNS = ["G", "H", "I", "J"]
DCF_COLUMNS = ["B", "C", "D", "E"]


def cell_value(sheet, ref):
    """The cached value of one cell, e.g. cell_value(wb["BS"], "H5")."""
    return sheet[ref].value


def check_dcf_labels(wb):
    """Stop if an inserted row has moved any cell the DCF rebuild reads."""
    for (sheet, ref), want in EXPECTED_LABELS.items():
        got = str(wb[sheet][ref].value or "").strip()
        if not got.startswith(want):
            raise ValueError(f"{sheet}!{ref} is {got!r}, expected a label starting {want!r}")


def single_value_check(label, derived, sheet, cell, tol=TOLERANCE):
    """A check with one number: the rebuilt figure against one cell of the workbook."""
    return Check(f"DCF {label}", [derived], [cell_value(sheet, cell)], ["DCF"], tol=tol)


def working_capital_change(bs):
    """Change in working capital per forecast year, from the balance-sheet lines.

    More assets tie up cash and more liabilities release it, so the asset
    movement is subtracted and the liability movement is added.
    """
    changes = []
    for i in range(len(BS_COLUMNS)):
        this_col = BS_COLUMNS[i]
        last_col = BS_PRIOR_COLUMNS[i]

        asset_movement = 0
        for row in WC_ASSETS:
            asset_movement += cell_value(bs, f"{this_col}{row}") - cell_value(bs, f"{last_col}{row}")

        liability_movement = 0
        for row in WC_LIABS:
            liability_movement += cell_value(bs, f"{this_col}{row}") - cell_value(bs, f"{last_col}{row}")

        changes.append(-asset_movement + liability_movement)
    return changes


def solve_implied_wacc(fcf, periods, growth, net_cash, shares, price_ref):
    """The WACC at which the DCF price equals the reference price, found by bisection.

    A higher WACC gives a lower price, so the answer sits between the lowest
    WACC the formula allows (just above g) and 50%.
    """
    low = growth + 1e-6
    high = 0.5
    price_at_low = _dcf_price(fcf, periods, low, growth, net_cash, shares)
    price_at_high = _dcf_price(fcf, periods, high, growth, net_cash, shares)
    if not price_at_low > price_ref > price_at_high:
        raise ValueError("the reference price is outside the range the reverse DCF can solve for")
    for _ in range(200):
        middle = (low + high) / 2
        if _dcf_price(fcf, periods, middle, growth, net_cash, shares) > price_ref:
            low = middle
        else:
            high = middle
    return (low + high) / 2


def build_dcf_checks(wb):
    """Rebuild working capital, the forecast balance sheet and the DCF from the statements."""
    IS, BS, CFS = wb["IS"], wb["BS"], wb["CFS"]
    A, D = wb["Assumptions"], wb["DCF"]
    check_dcf_labels(wb)

    checks = []

    # ---- working capital from the balance sheet, not from the cash flow statement
    wc_change = working_capital_change(BS)
    workbook_wc_change = [cell_value(CFS, f"{col}13") for col in STATEMENT_COLUMNS]
    checks.append(Check("DCF change in working capital = movement in the balance-sheet lines",
                        wc_change, workbook_wc_change, FORECAST_YEARS))

    # ---- the forecast balance sheet, re-added from its line items (no totals, no plug)
    gaps = []
    for col in BS_COLUMNS:
        assets = sum(cell_value(BS, f"{col}{row}") for row in BS_ASSETS)
        liabilities_and_equity = sum(cell_value(BS, f"{col}{row}") for row in BS_LIABS_EQUITY)
        gaps.append(assets - liabilities_and_equity)
    checks.append(Check("DCF forecast balance sheet balances from its line items",
                        gaps, [0.0] * 4, FORECAST_YEARS))

    cfs_ending_cash = [cell_value(CFS, f"{col}30") for col in STATEMENT_COLUMNS]
    bs_cash = [cell_value(BS, f"{col}5") for col in BS_COLUMNS]
    checks.append(Check("DCF forecast cash = cash flow statement ending cash",
                        cfs_ending_cash, bs_cash, FORECAST_YEARS))

    # ---- WACC from the Assumptions tab inputs
    cost_of_equity = (cell_value(A, RISK_FREE_RATE)
                      + cell_value(A, BETA) * cell_value(A, EQUITY_RISK_PREMIUM))
    debt_weight = cell_value(A, TARGET_DEBT_TO_CAPITAL)
    after_tax_cost_of_debt = cell_value(A, PRE_TAX_COST_OF_DEBT) * (1 - cell_value(A, WACC_TAX_RATE))
    wacc = cost_of_equity * (1 - debt_weight) + after_tax_cost_of_debt * debt_weight
    growth = cell_value(A, TERMINAL_GROWTH)

    # ---- unlevered free cash flow: NOPAT + D&A - CapEx - change in working capital
    # CFS capex is negative, and the working-capital change rebuilt above from the
    # balance sheet is already in cash-flow sign (more assets is negative), so both
    # are added.
    fcf = []
    for i in range(len(STATEMENT_COLUMNS)):
        col = STATEMENT_COLUMNS[i]
        tax_rate = cell_value(A, f"{col}25")
        nopat = cell_value(IS, f"{col}28") * (1 - tax_rate)
        fcf.append(nopat + cell_value(CFS, f"{col}10") + cell_value(CFS, f"{col}19") + wc_change[i])

    # ---- discounting from the valuation date to each fiscal year-end
    valuation_date = cell_value(A, VALUATION_DATE)
    periods = []
    for col in DCF_COLUMNS:
        year_end = cell_value(D, f"{col}14")
        periods.append((year_end - valuation_date).days / 365.25)

    pv_of_cash_flows = 0
    for cash_flow, t in zip(fcf, periods):
        if t > 0:  # a year that ended before the valuation date is already in net cash
            pv_of_cash_flows += cash_flow / (1 + wacc) ** t
    pv_of_terminal = fcf[-1] * (1 + growth) / (wacc - growth) / (1 + wacc) ** periods[-1]

    # ---- net cash at FY26E, the last balance sheet that closed before the valuation date
    last_closed = BS_COLUMNS[0]            # column H
    total_debt = (cell_value(BS, f"{last_closed}24") + cell_value(BS, f"{last_closed}25")
                  + cell_value(BS, f"{last_closed}29"))
    net_cash = (cell_value(BS, f"{last_closed}5") + cell_value(BS, f"{last_closed}6")
                + cell_value(BS, f"{last_closed}15") - total_debt)

    shares = cell_value(A, DILUTED_SHARES)
    price_ref = cell_value(A, REFERENCE_PRICE)
    enterprise_value = pv_of_cash_flows + pv_of_terminal
    price = _dcf_price(fcf, periods, wacc, growth, net_cash, shares)
    discount_factor_last = 1 / (1 + wacc) ** periods[-1]
    final_year_col = STATEMENT_COLUMNS[-1]                     # column J, FY29E
    ebitda_fy29 = cell_value(IS, f"{final_year_col}28") + cell_value(CFS, f"{final_year_col}10")

    years_included = []
    for year, t in zip(FORECAST_YEARS, periods):
        if t > 0:
            years_included.append(year)

    # ---- compare each rebuilt figure with the DCF tab
    checks.append(single_value_check(
        "WACC = cost of equity x E/V + after-tax cost of debt x D/V", wacc, A, WACC_CELL, tol=1e-9))

    for i in range(len(FORECAST_YEARS)):
        checks.append(single_value_check(
            f"unlevered free cash flow {FORECAST_YEARS[i]} = NOPAT at the scenario tax rate + D&A - CapEx - change in WC",
            fcf[i], D, f"{DCF_COLUMNS[i]}13"))

    for i in range(len(FORECAST_YEARS)):
        checks.append(single_value_check(
            f"years from valuation date to {FORECAST_YEARS[i]} year-end",
            periods[i], D, f"{DCF_COLUMNS[i]}15", tol=1e-9))

    checks.append(single_value_check(
        f"sum of discounted cash flows ({', '.join(years_included)})", pv_of_cash_flows, D, "B20"))
    checks.append(single_value_check(
        "terminal value, discounted from the FY29E year-end", pv_of_terminal, D, "B21"))
    checks.append(single_value_check("enterprise value", enterprise_value, D, "B22"))
    checks.append(single_value_check(
        "net cash (FY26E) = cash + marketable securities - total debt", net_cash, D, "B23"))
    checks.append(single_value_check("implied share price", price, D, "B26", tol=0.01))
    checks.append(single_value_check(
        "upside = implied price / reference price - 1", price / price_ref - 1, D, "B28", tol=1e-6))
    checks.append(single_value_check(
        "sensitivity grid centre equals the headline price", price, D, "D36", tol=0.01))

    corner_wacc_up = _dcf_price(fcf, periods, wacc + 0.005, growth - 0.005, net_cash, shares)
    checks.append(single_value_check(
        "grid corner, WACC +50 bps and g -50 bps", corner_wacc_up, D, "C37", tol=0.01))
    corner_wacc_down = _dcf_price(fcf, periods, wacc - 0.005, growth + 0.005, net_cash, shares)
    checks.append(single_value_check(
        "grid corner, WACC -50 bps and g +50 bps", corner_wacc_down, D, "E35", tol=0.01))

    pe_bar = cell_value(IS, f"{final_year_col}43") * cell_value(A, PE_MID_MULTIPLE) * discount_factor_last
    checks.append(single_value_check(
        "P/E bar (mid multiple x FY29E diluted EPS, discounted)", pe_bar, D, "C69", tol=0.01))
    ev_ebitda_bar = (ebitda_fy29 * cell_value(A, EV_EBITDA_MID_MULTIPLE) * discount_factor_last + net_cash) / shares
    checks.append(single_value_check(
        "EV/EBITDA bar (mid multiple x FY29E EBITDA, discounted, plus net cash)",
        ev_ebitda_bar, D, "C70", tol=0.01))

    # ---- the summary figures printed at the end of the report
    implied_wacc = solve_implied_wacc(fcf, periods, growth, net_cash, shares, price_ref)
    info = {"price": price, "price_ref": price_ref, "wacc": wacc, "g": growth,
            "implied_wacc": implied_wacc, "tv_share": pv_of_terminal / enterprise_value,
            "scenario": cell_value(A, ACTIVE_SCENARIO), "ev": enterprise_value,
            "net_cash": net_cash, "equity_value": enterprise_value + net_cash,
            "market_value": price_ref * shares,
            "bear": cell_value(D, "B68"), "bull": cell_value(D, "D68")}
    return checks, info


def main(argv):
    """Load the workbook, run every check group, print the report and return the exit code (0 when all pass)."""
    path = Path(argv[1]) if len(argv) > 1 else DEFAULT_MODEL
    if not path.exists():
        sys.exit(f"workbook not found: {path}")

    wb = openpyxl.load_workbook(path, data_only=True)
    if "Validation" not in wb.sheetnames:
        sys.exit(f"no Validation sheet in {path.name}; sheets are {wb.sheetnames}")

    ws = wb["Validation"]

    # Verified, not assumed: same years across all three statements, and no
    # duplicate row label that a check could silently bind to the wrong row.
    years = check_year_headers(ws, find_section_headers(ws))
    rows = read_rows(ws)

    # A second load, formulas rather than values, purely to read the source-cell
    # mapping out of the Validation tab's column G.
    wb_formulas = openpyxl.load_workbook(path, data_only=False)
    links = read_source_links(wb_formulas["Validation"])
    if not links:
        sys.exit(
            "No source links found in the Validation tab's column G. Without them "
            "this script can only check the transcription against itself, which is "
            "the failure mode it exists to avoid. Refusing to report a pass."
        )

    print(f"Model:  {path.name}")
    print(f"Years:  {', '.join(str(y) for y in years)}")
    print(f"Tolerance: {TOLERANCE:,.0f} (millions USD, for the workbook's rounding)")
    print("-" * 72)

    dcf_checks, dcf = build_dcf_checks(wb)
    ratio_checks = build_ratio_checks(wb, rows, years)
    checks = build_checks(rows, years) + build_transcription_checks(wb, links, years) + dcf_checks + ratio_checks
    failed = 0
    for check in checks:
        if check.passed:
            print(f"  PASS  {check.label}")
        elif check.advisory:
            print(f"  NOTE  {check.label}")
        else:
            failed += 1
            print(f"  FAIL  {check.label}")
            for year, exp, act, why in check.failures:
                print(f"          {year}: derived {exp!r} vs workbook {act!r} ({why})")
            if check.note:
                print()
                for line in _wrap(check.note, 66):
                    print(f"          {line}")
                print()

    hard = [c for c in checks if not c.advisory]
    notes = [c for c in checks if c.advisory and not c.passed]
    identities = [c for c in hard if not c.label.startswith(("SRC ", "DCF ", "RAT "))]
    ratios = [c for c in hard if c.label.startswith("RAT ")]
    ratio_failures = sum(1 for c in ratios if not c.passed)
    sources = [c for c in hard if c.label.startswith("SRC ")]
    valuation = [c for c in hard if c.label.startswith("DCF ")]
    valuation_failures = sum(1 for c in valuation if not c.passed)
    identity_failures = sum(1 for c in identities if not c.passed)
    source_failures = sum(1 for c in sources if not c.passed)

    print("-" * 72)
    print(f"{len(identities) - identity_failures} of {len(identities)} accounting "
          f"identities re-derived independently and matched.")
    print(f"{len(sources) - source_failures} of {len(sources)} Validation-tab line "
          f"items match the IS / BS / CFS cells they were transcribed from, "
          f"across {len(years)} years each.")
    print(f"{len(valuation) - valuation_failures} of {len(valuation)} valuation figures re-derived "
          f"from the forecast statements and inputs ({dcf['scenario']} scenario).")
    print(f"{len(ratios) - ratio_failures} of {len(ratios)} Ratios-tab rows rebuilt from the traced line items "
          f"and matched in every year.")
    print(f"DCF: ${dcf['price']:,.2f} a share at WACC {dcf['wacc']:.2%} and g {dcf['g']:.1%}, against "
          f"${dcf['price_ref']:,.2f}. The reference price implies a WACC of {dcf['implied_wacc']:.2%} "
          f"on the same cash flows. Terminal value is {dcf['tv_share']:.0%} of enterprise value.")
    print(f"Enterprise value ${dcf['ev'] / 1e6:,.2f}tn plus FY26E net cash ${dcf['net_cash'] / 1e3:,.1f}bn.")
    print(f"Bear / Bull prices as recorded in the workbook (not re-derived): "
          f"${dcf['bear']:,.2f} / ${dcf['bull']:,.2f}.")
    print(f"Equity value ${dcf['equity_value'] / 1e6:,.2f}tn against a market value of "
          f"${dcf['market_value'] / 1e6:,.2f}tn at the reference price: the market pays "
          f"${(dcf['market_value'] - dcf['equity_value']) / 1e6:,.2f}tn more than the base case supports.")
    if notes:
        print(f"{len(notes)} advisory check(s) flagged above: explained presentation "
              f"differences, not arithmetic. See the note under each.")
    if failed:
        print(f"{failed} FAILED. These are arithmetic, and the workbook's own "
              f"Validation tab reports a tick for them anyway.")
        return 1
    print("Every total was recomputed outside the spreadsheet and agrees, every line "
          "item traces to the statement it came from, and the valuation rebuilds.")
    print("Not covered: the Dashboard and Pivots tabs; the Bear and Bull prices are "
          "printed as recorded.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
