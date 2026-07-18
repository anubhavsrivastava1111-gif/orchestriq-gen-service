"""
OrchestrIQ Excel Engine — CFO-Grade Financial Models via openpyxl
Architecture: AI extracts parameters → Python builds deterministically.
Every formula is hardcoded. No AI hallucination in cell references.
"""

import io
from datetime import datetime, date
from typing import Any, Dict, List, Optional

from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Border, Side, Alignment, numbers
)
from openpyxl.styles.numbers import FORMAT_NUMBER_COMMA_SEPARATED1
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.worksheet.datavalidation import DataValidation

# ─── BRAND PALETTE ────────────────────────────────────────────────────────────
NAVY        = "0D1B2A"
NAVY_MID    = "1A2744"
SLATE       = "2C3E50"
GOLD        = "C9A84C"
GOLD_LT     = "F5E6A3"
WHITE       = "FFFFFF"
LIGHT_GR    = "F7F8FA"
MID_GR      = "D6D8DB"
DARK_GR     = "4A4A4A"
GREEN_POS   = "1A7A4A"
GREEN_LT    = "E6F4ED"
RED_NEG     = "C0392B"
RED_LT      = "FDE8E6"
AMBER       = "D68910"
AMBER_LT    = "FEF9E7"
BLUE_IN     = "1F4E79"
YELLOW_IN   = "FFFDE7"   # INPUT CELLS — yellow background, blue text
TOTAL_BG    = "E8EBF0"   # Subtotal row background
GRAND_BG    = "D0D7E2"   # Grand total row

# ─── HELPERS ──────────────────────────────────────────────────────────────────

def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, size=10, color=DARK_GR, italic=False,
          name="Calibri") -> Font:
    return Font(name=name, bold=bold, size=size,
                color=color, italic=italic)

def _side(style="thin", color=MID_GR) -> Side:
    return Side(style=style, color=color)

def _border(style="thin", color=MID_GR) -> Border:
    s = _side(style, color)
    return Border(left=s, right=s, top=s, bottom=s)

def _bottom(style="medium", color=GOLD) -> Border:
    return Border(bottom=Side(style=style, color=color))

def _top_bottom(top_color=GOLD, bot_color=GOLD, style="medium") -> Border:
    return Border(top=Side(style=style, color=top_color),
                  bottom=Side(style=style, color=bot_color))

def _align(h="left", v="center", wrap=False) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _num_fmt(sym: str, decimals=0) -> str:
    """Build Excel number format string."""
    if sym in ("₹", "Rs", "INR"):
        d = "0" * decimals
        if decimals:
            return f'[₹]#,##0.{d}_);([₹]#,##0.{d})'
        return '[₹]#,##0_);([₹]#,##0)'
    if sym == "$":
        return f'"$"#,##0{"." + "0"*decimals if decimals else ""}_);("$"#,##0)'
    return f'#,##0{"." + "0"*decimals if decimals else ""}_);(#,##0)'

def _pct_fmt(decimals=1) -> str:
    return f'0.{"0"*decimals}%'

def _x_fmt(decimals=1) -> str:
    return f'0.{"0"*decimals}"x"'

def _num(v) -> float:
    try: return float(str(v or 0).replace(",", "").replace("%", ""))
    except: return 0.0

def _col(n: int) -> str:
    """Column number → letter (1-based)."""
    return get_column_letter(n)

# ─── STYLER CLASS ─────────────────────────────────────────────────────────────

class Styler:
    """Apply consistent styling to cells."""
    def __init__(self, ws, sym="₹"):
        self.ws  = ws
        self.sym = sym

    def cell(self, row, col):
        return self.ws.cell(row=row, column=col)

    def header_row(self, row, cols, texts, merge_end=None):
        """Full-width navy header row."""
        for c, txt in zip(cols, texts):
            cell = self.ws.cell(row=row, column=c, value=txt)
            cell.fill      = _fill(NAVY)
            cell.font      = _font(bold=True, size=9.5, color=WHITE)
            cell.alignment = _align("center")
            cell.border    = _border("thin", NAVY)

    def section_title(self, row, col, text, span=None):
        """Navy section title, optionally merged."""
        cell = self.ws.cell(row=row, column=col, value=text)
        cell.fill      = _fill(NAVY_MID)
        cell.font      = _font(bold=True, size=10, color=WHITE)
        cell.alignment = _align("left")
        if span:
            self.ws.merge_cells(
                start_row=row, start_column=col,
                end_row=row, end_column=col + span - 1)

    def label(self, row, col, text, indent=0, bold=False, italic=False):
        """Row label."""
        cell = self.ws.cell(row=row, column=col, value=" " * (indent * 2) + text)
        cell.font      = _font(bold=bold, size=9, color=DARK_GR, italic=italic)
        cell.alignment = _align("left")
        cell.border    = _border("thin", MID_GR)
        return cell

    def input_cell(self, row, col, value=None, fmt=None):
        """Yellow input cell — user-editable assumption."""
        cell = self.ws.cell(row=row, column=col, value=value)
        cell.fill      = _fill(YELLOW_IN)
        cell.font      = _font(bold=False, size=9, color=BLUE_IN)
        cell.alignment = _align("right")
        cell.border    = Border(
            left=Side(style="thin", color=BLUE_IN),
            right=Side(style="thin", color=BLUE_IN),
            top=Side(style="thin", color=BLUE_IN),
            bottom=Side(style="thin", color=BLUE_IN),
        )
        if fmt: cell.number_format = fmt
        return cell

    def formula_cell(self, row, col, formula, fmt=None, bold=False,
                     fill=None, color=DARK_GR):
        """Formula cell — calculated from inputs."""
        cell = self.ws.cell(row=row, column=col, value=formula)
        cell.fill      = _fill(fill or WHITE)
        cell.font      = _font(bold=bold, size=9, color=color)
        cell.alignment = _align("right")
        cell.border    = _border("thin", MID_GR)
        if fmt: cell.number_format = fmt
        return cell

    def total_cell(self, row, col, formula, fmt=None, level="sub"):
        """Subtotal or grand-total cell."""
        bg  = GRAND_BG if level == "grand" else TOTAL_BG
        col_hex = NAVY   if level == "grand" else SLATE
        cell = self.ws.cell(row=row, column=col, value=formula)
        cell.fill      = _fill(bg)
        cell.font      = _font(bold=True, size=9, color=col_hex)
        cell.alignment = _align("right")
        cell.border    = _top_bottom(GOLD, GOLD, "medium")
        if fmt: cell.number_format = fmt
        return cell

    def pct_cell(self, row, col, formula, bold=False):
        """Percentage cell."""
        cell = self.ws.cell(row=row, column=col, value=formula)
        cell.fill         = _fill(LIGHT_GR)
        cell.font         = _font(bold=bold, size=9, color=DARK_GR)
        cell.alignment    = _align("right")
        cell.border       = _border("thin", MID_GR)
        cell.number_format = _pct_fmt()
        return cell

    def green_cell(self, row, col, formula, fmt=None):
        cell = self.ws.cell(row=row, column=col, value=formula)
        cell.fill      = _fill(GREEN_LT)
        cell.font      = _font(bold=True, size=9, color=GREEN_POS)
        cell.alignment = _align("right")
        cell.border    = _border("thin", MID_GR)
        if fmt: cell.number_format = fmt
        return cell

    def red_cell(self, row, col, formula, fmt=None):
        cell = self.ws.cell(row=row, column=col, value=formula)
        cell.fill      = _fill(RED_LT)
        cell.font      = _font(bold=True, size=9, color=RED_NEG)
        cell.alignment = _align("right")
        cell.border    = _border("thin", MID_GR)
        if fmt: cell.number_format = fmt
        return cell

    def set_col_widths(self, widths: Dict[int, float]):
        for col, w in widths.items():
            self.ws.column_dimensions[_col(col)].width = w

    def set_row_height(self, row, height=16):
        self.ws.row_dimensions[row].height = height


# ─── PERIOD LABELS ────────────────────────────────────────────────────────────

MONTHS = ["Jan","Feb","Mar","Apr","May","Jun",
          "Jul","Aug","Sep","Oct","Nov","Dec"]

def _periods(start_month, num: int) -> List[str]:
    """Generate period labels: 'Jan-25', 'Feb-25', …"""
    # Accept int (1-12), string like "Apr-25" or "4", or None
    if isinstance(start_month, int):
        mi = max(0, min(start_month - 1, 11))
        yr = 25
        labels = []
        for i in range(num):
            labels.append(f"{MONTHS[(mi + i) % 12]}-{str(yr + (mi + i) // 12)[-2:]}")
        return labels
    sm = str(start_month).strip() if start_month else "Apr-25"
    try:
        if sm.isdigit():
            mi = max(0, min(int(sm) - 1, 11))
            yr = 25
        else:
            m_name = sm[:3].capitalize()
            yr = int(sm.split("-")[-1]) if "-" in sm else 25
            mi = MONTHS.index(m_name)
    except (ValueError, IndexError):
        mi, yr = 3, 25  # Default: Apr-25
    labels = []
    for i in range(num):
        labels.append(f"{MONTHS[(mi + i) % 12]}-{str(yr + (mi + i) // 12)[-2:]}")
    return labels

def _year_labels(periods: List[str]) -> List[str]:
    """Derive annual column labels from monthly periods."""
    years = {}
    for p in periods:
        y = p.split("-")[-1]
        years[y] = years.get(y, 0) + 1
    return [f"FY{y}" for y in years.keys()]

# ─── INSTRUCTIONS SHEET ───────────────────────────────────────────────────────

def _add_instructions(wb: Workbook, params: Dict) -> None:
    ws = wb.create_sheet("📋 Instructions")
    ws.sheet_properties.tabColor = "808080"
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 65

    def row(r, a, b, fill=None, bold=False):
        ca = ws.cell(row=r, column=1, value=a)
        cb = ws.cell(row=r, column=2, value=b)
        if fill:
            ca.fill = cb.fill = _fill(fill)
        ca.font = cb.font = _font(bold=bold, size=9)
        ca.alignment = _align("left")
        cb.alignment = _align("left", wrap=True)
        ws.row_dimensions[r].height = 18

    ws.merge_cells("A1:B1")
    c = ws.cell(row=1, column=1,
                value="OrchestrIQ Financial Model — Usage Guide")
    c.fill = _fill(NAVY); c.font = _font(bold=True, size=13, color=WHITE)
    c.alignment = _align("center")
    ws.row_dimensions[1].height = 30

    guide = [
        ("", ""),
        ("COLOUR CODING", "", NAVY_MID, True),
        ("🟡 Yellow (Blue text)", "INPUT CELLS — Enter your assumptions here only", YELLOW_IN),
        ("⚪ White", "FORMULA CELLS — Auto-calculated. Do not edit.", WHITE),
        ("🔵 Light Blue / Grey", "TOTAL / SUBTOTAL rows — Summarised automatically", TOTAL_BG),
        ("🟢 Green", "Positive variance or KPI indicator", GREEN_LT),
        ("🔴 Red / Pink", "Negative variance or risk indicator", RED_LT),
        ("", ""),
        ("WORKFLOW", "", NAVY_MID, True),
        ("Step 1", "Go to 📊 Assumptions tab. Fill all yellow cells.", YELLOW_IN),
        ("Step 2", "Verify P&L sheet updates automatically."),
        ("Step 3", "Check Cash Flow for runway position."),
        ("Step 4", "Review Dashboard for executive summary."),
        ("", ""),
        ("RULES", "", NAVY_MID, True),
        ("❌ Never edit", "Formula cells (white/grey cells)"),
        ("❌ Never delete", "Row or column headers"),
        ("✅ Only edit", "Yellow cells in Assumptions sheet"),
        ("", ""),
        ("GENERATED BY", "", NAVY_MID, True),
        ("Platform", "OrchestrIQ | GorakhAI"),
        ("Company", params.get("company_name", "—")),
        ("Model Type", params.get("template_type", "financial_dashboard").replace("_", " ").title()),
        ("Generated", datetime.now().strftime("%d %b %Y %H:%M")),
        ("Currency", params.get("currency", "INR") + " (" + params.get("currency_symbol", "₹") + ")"),
    ]
    for i, item in enumerate(guide, 2):
        if len(item) == 4:
            row(i, item[0], item[1], item[2], item[3])
        elif len(item) == 3:
            row(i, item[0], item[1], item[2])
        else:
            row(i, item[0], item[1] if len(item) > 1 else "")

# ─── ASSUMPTIONS SHEET ────────────────────────────────────────────────────────

def _add_assumptions(wb: Workbook, params: Dict,
                     periods: List[str]) -> None:
    ws = wb.create_sheet("📊 Assumptions")
    ws.sheet_properties.tabColor = GOLD
    st = Styler(ws, params.get("currency_symbol", "₹"))
    sym = params.get("currency_symbol", "₹")
    nf  = _num_fmt(sym)

    # Freeze header
    ws.freeze_panes = "B3"
    st.set_col_widths({1: 34, 2: 20, 3: 25})

    # Title
    ws.merge_cells("A1:C1")
    c = ws.cell(row=1, column=1,
                value=f"{params.get('company_name','Company')} — Financial Assumptions")
    c.fill = _fill(NAVY); c.font = _font(bold=True, size=13, color=WHITE)
    c.alignment = _align("center")
    ws.row_dimensions[1].height = 28

    # Column headers
    for col, hdr in enumerate(["Assumption", "Value", "Notes"], 1):
        cell = ws.cell(row=2, column=col, value=hdr)
        cell.fill = _fill(SLATE)
        cell.font = _font(bold=True, size=9, color=WHITE)
        cell.alignment = _align("center")
    ws.row_dimensions[2].height = 16

    A = params.get("assumptions", {})
    r = 3

    def section(title):
        nonlocal r
        ws.merge_cells(f"A{r}:C{r}")
        cell = ws.cell(row=r, column=1, value=f"  {title}")
        cell.fill = _fill(NAVY_MID); cell.font = _font(bold=True, size=9, color=GOLD_LT)
        ws.row_dimensions[r].height = 16
        r += 1

    def inp(label, key, default, fmt=None, note=""):
        nonlocal r
        st.label(r, 1, label)
        v = _num(A.get(key, default))
        st.input_cell(r, 2, v, fmt or nf)
        ws.cell(row=r, column=3, value=note).font = _font(size=8.5, italic=True, color=DARK_GR)
        ws.row_dimensions[r].height = 15
        r += 1

    def text_inp(label, key, default, note=""):
        nonlocal r
        st.label(r, 1, label)
        cell = ws.cell(row=r, column=2, value=A.get(key, default))
        cell.fill = _fill(YELLOW_IN); cell.font = _font(size=9, color=BLUE_IN)
        cell.alignment = _align("left")
        ws.cell(row=r, column=3, value=note).font = _font(size=8.5, italic=True, color=DARK_GR)
        ws.row_dimensions[r].height = 15
        r += 1

    # ── Company & Model ───────────────────────────────────────────────────────
    section("COMPANY & MODEL PARAMETERS")
    text_inp("Company Name",       "company_name",  params.get("company_name","Company"))
    text_inp("Financial Year Start","start_month",  params.get("start_month","Apr-25"),  "e.g. Apr-25")
    text_inp("Reporting Currency", "currency",      params.get("currency","INR"))
    inp("Model Periods (months)",  "num_periods",   params.get("num_periods",12), "0", "12, 24, or 36")
    r += 1

    # ── Revenue ───────────────────────────────────────────────────────────────
    section("REVENUE ASSUMPTIONS")
    inp("Base Monthly Revenue",    "base_revenue",  A.get("base_revenue", 5000000), nf, "Month 1 top-line")
    inp("Monthly Revenue Growth (%)", "rev_growth", A.get("rev_growth", 5), _pct_fmt(), "MoM % increase")
    inp("Revenue Stream 2 (%)",    "rev_stream2",   A.get("rev_stream2", 20), _pct_fmt(), "% of total from stream 2")
    inp("Revenue Stream 3 (%)",    "rev_stream3",   A.get("rev_stream3", 10), _pct_fmt(), "% of total from stream 3")
    inp("Seasonality Factor Q4 (%)", "seasonality_q4", A.get("seasonality_q4", 15), _pct_fmt(), "Uplift in Q4 months")
    r += 1

    # ── Cost of Revenue ───────────────────────────────────────────────────────
    section("COST OF REVENUE")
    inp("COGS as % of Revenue",    "cogs_pct",      A.get("cogs_pct", 35), _pct_fmt(), "Direct variable costs")
    inp("Gross Margin Target (%)", "gm_target",     A.get("gm_target", 65), _pct_fmt(), "Benchmark target")
    r += 1

    # ── Operating Expenses ────────────────────────────────────────────────────
    section("OPERATING EXPENSES (MONTHLY BASE)")
    inp("Salaries & Benefits",     "salary_base",   A.get("salary_base", 1500000), nf, "Total monthly payroll")
    inp("Annual Salary Increase (%)", "salary_growth", A.get("salary_growth", 10), _pct_fmt(), "YoY increase")
    inp("Marketing & Advertising", "marketing",     A.get("marketing", 300000), nf)
    inp("Technology / SaaS / Cloud","technology",   A.get("technology", 150000), nf)
    inp("Rent & Facilities",       "rent",          A.get("rent", 200000), nf)
    inp("General & Administrative","ga",            A.get("ga", 250000), nf)
    inp("Professional Fees",       "prof_fees",     A.get("prof_fees", 100000), nf)
    inp("Other Operating Expenses","other_opex",    A.get("other_opex", 100000), nf)
    r += 1

    # ── Non-Operating ─────────────────────────────────────────────────────────
    section("NON-OPERATING & TAX")
    inp("Depreciation (Monthly)",  "depreciation",  A.get("depreciation", 50000), nf)
    inp("Interest Income / (Expense)", "interest",  A.get("interest", 0), nf, "Positive = income")
    inp("Tax Rate (%)",            "tax_rate",      A.get("tax_rate", 25), _pct_fmt())
    r += 1

    # ── Balance Sheet / Cash ──────────────────────────────────────────────────
    section("BALANCE SHEET & CASH FLOW")
    inp("Opening Cash Balance",    "opening_cash",  A.get("opening_cash", 10000000), nf)
    inp("Accounts Receivable Days","ar_days",       A.get("ar_days", 45), "0", "DSO — collection days")
    inp("Accounts Payable Days",   "ap_days",       A.get("ap_days", 30), "0", "DPO — payment days")
    inp("Capex (Monthly Average)", "capex",         A.get("capex", 100000), nf)
    inp("Working Capital Reserve", "wc_reserve",    A.get("wc_reserve", 2000000), nf)
    r += 1

    # Legend
    ws.merge_cells(f"A{r+1}:C{r+1}")
    c = ws.cell(row=r+1, column=1,
                value="🟡 Yellow = Input Cell (edit here)   ⚪ White = Formula (do not edit)")
    c.font = _font(size=8.5, italic=True, color=BLUE_IN)
    c.fill = _fill(GOLD_LT)

# ─── P&L SHEET ────────────────────────────────────────────────────────────────

def _add_pl(wb: Workbook, params: Dict, periods: List[str]) -> None:
    ws = wb.create_sheet("📈 P&L")
    ws.sheet_properties.tabColor = NAVY
    st  = Styler(ws, params.get("currency_symbol", "₹"))
    sym = params.get("currency_symbol", "₹")
    nf  = _num_fmt(sym)
    pf  = _pct_fmt()
    A   = params.get("assumptions", {})
    N   = len(periods)
    # Column layout: col 1 = label, col 2..N+1 = months, col N+2.. = annual summaries
    C_FIRST = 2   # first data column (1-indexed)
    C_LAST  = N + 1

    ws.freeze_panes = f"B4"
    st.set_col_widths({1: 32})
    for c in range(C_FIRST, C_LAST + 6):
        ws.column_dimensions[_col(c)].width = 13

    # ── Title row ─────────────────────────────────────────────────────────────
    ws.merge_cells(f"A1:{_col(C_LAST+4)}1")
    c1 = ws.cell(row=1, column=1,
                 value=f"{params.get('company_name','Company')} — Profit & Loss Statement")
    c1.fill = _fill(NAVY); c1.font = _font(bold=True, size=12, color=WHITE)
    c1.alignment = _align("center")
    ws.row_dimensions[1].height = 26

    # Subtitle row
    ws.merge_cells(f"A2:{_col(C_LAST+4)}2")
    c2 = ws.cell(row=2, column=1,
                 value=f"All figures in {params.get('currency','INR')} | Monthly View + Annual Summary | Formulas auto-calculated from Assumptions sheet")
    c2.fill = _fill(SLATE); c2.font = _font(size=8.5, italic=True, color=GOLD_LT)
    c2.alignment = _align("center")

    # ── Header row (periods) ──────────────────────────────────────────────────
    ws.cell(row=3, column=1, value="Line Item").fill = _fill(NAVY)
    ws.cell(row=3, column=1).font = _font(bold=True, size=9, color=WHITE)
    ws.cell(row=3, column=1).alignment = _align("center")

    for i, p in enumerate(periods):
        c = ws.cell(row=3, column=C_FIRST + i, value=p)
        c.fill = _fill(NAVY); c.font = _font(bold=True, size=8.5, color=WHITE)
        c.alignment = _align("center")

    # Annual summary columns
    years = _year_labels(periods)
    n_years = len(years)
    yr_start_col = C_LAST + 2
    for j, yr in enumerate(years):
        c = ws.cell(row=3, column=yr_start_col + j, value=yr)
        c.fill = _fill(GOLD); c.font = _font(bold=True, size=9, color=NAVY)
        c.alignment = _align("center")

    ws.row_dimensions[3].height = 16

    # ── Assumption references (hidden reference row) ───────────────────────────
    # A!B2 = base_revenue, A!B3 = rev_growth, A!B4 = rev_stream2, etc.
    # We reference Assumptions sheet cells directly:
    ASM = "='📊 Assumptions'!"  # prefix for cross-sheet refs

    # Map assumption row numbers on Assumptions sheet
    # Row 3=company, 4=start_month, 5=currency, 6=periods, 8=base_revenue,
    # 9=rev_growth, 10=stream2, 11=stream3, 12=seasonality
    # 14=cogs_pct, 15=gm_target
    # 17=salary, 18=sal_growth, 19=marketing, 20=tech, 21=rent, 22=ga, 23=prof, 24=other
    # 26=depr, 27=interest, 28=tax
    # 30=cash, 31=ar, 32=ap, 33=capex, 34=wc

    BASE_REV   = f"{ASM}B8"
    REV_GR     = f"{ASM}B9"
    S2_PCT     = f"{ASM}B10"
    S3_PCT     = f"{ASM}B11"
    SEAS_Q4    = f"{ASM}B12"
    COGS_PCT   = f"{ASM}B14"
    SAL_BASE   = f"{ASM}B17"
    SAL_GR     = f"{ASM}B18"
    MKTG       = f"{ASM}B19"
    TECH       = f"{ASM}B20"
    RENT       = f"{ASM}B21"
    GA         = f"{ASM}B22"
    PROF       = f"{ASM}B23"
    OTHER_OPX  = f"{ASM}B24"
    DEPR       = f"{ASM}B26"
    INTEREST   = f"{ASM}B27"
    TAX_RATE   = f"{ASM}B28"

    # ── Helper: get cell address for P&L row r, month i ──────────────────────
    def addr(row_r, month_i):
        return f"{_col(C_FIRST + month_i)}{row_r}"

    def rng(row_r, i_start=0, i_end=None):
        if i_end is None: i_end = N - 1
        return f"{_col(C_FIRST+i_start)}{row_r}:{_col(C_FIRST+i_end)}{row_r}"

    row = 4  # current write row

    def sep():
        nonlocal row
        ws.row_dimensions[row].height = 4
        row += 1

    # ── REVENUE ───────────────────────────────────────────────────────────────
    st.section_title(row, 1, "REVENUE", C_LAST + n_years + 1)
    ws.row_dimensions[row].height = 16
    row += 1

    R_REV1 = row   # Revenue Stream 1 row
    for i in range(N):
        col = C_FIRST + i
        if i == 0:
            formula = f"={BASE_REV}*(1-{S2_PCT}-{S3_PCT})"
        else:
            prev    = _col(col - 1)
            formula = f"={prev}{R_REV1}*(1+{REV_GR})"
        st.formula_cell(row, col, formula, nf)
    st.label(row, 1, "  Revenue — Core Product", bold=False)
    # Annuals
    months_per_yr = N // max(1, n_years)
    for j in range(n_years):
        c_start = C_FIRST + j * months_per_yr
        c_end   = C_FIRST + min((j+1)*months_per_yr, N) - 1
        st.total_cell(row, yr_start_col+j,
                      f"=SUM({_col(c_start)}{row}:{_col(c_end)}{row})",
                      nf, "sub")
    row += 1

    R_REV2 = row   # Revenue Stream 2
    for i in range(N):
        col = C_FIRST + i
        formula = f"={_col(col)}{R_REV1}*{S2_PCT}/(1-{S2_PCT}-{S3_PCT})"
        st.formula_cell(row, col, formula, nf)
    st.label(row, 1, "  Revenue — Stream 2", bold=False)
    for j in range(n_years):
        c_start = C_FIRST + j * months_per_yr
        c_end   = C_FIRST + min((j+1)*months_per_yr, N) - 1
        st.total_cell(row, yr_start_col+j,
                      f"=SUM({_col(c_start)}{row}:{_col(c_end)}{row})",
                      nf, "sub")
    row += 1

    R_REV3 = row   # Revenue Stream 3
    for i in range(N):
        col = C_FIRST + i
        formula = f"={_col(col)}{R_REV1}*{S3_PCT}/(1-{S2_PCT}-{S3_PCT})"
        st.formula_cell(row, col, formula, nf)
    st.label(row, 1, "  Revenue — Stream 3", bold=False)
    for j in range(n_years):
        c_start = C_FIRST + j * months_per_yr
        c_end   = C_FIRST + min((j+1)*months_per_yr, N) - 1
        st.total_cell(row, yr_start_col+j,
                      f"=SUM({_col(c_start)}{row}:{_col(c_end)}{row})",
                      nf, "sub")
    row += 1

    R_TOT_REV = row  # ← TOTAL REVENUE
    for i in range(N):
        col = C_FIRST + i
        st.total_cell(row, col,
                      f"={_col(col)}{R_REV1}+{_col(col)}{R_REV2}+{_col(col)}{R_REV3}",
                      nf, "grand")
    st.label(row, 1, "TOTAL REVENUE", bold=True)
    ws.cell(row=row, column=1).fill = _fill(TOTAL_BG)
    for j in range(n_years):
        c_start = C_FIRST + j * months_per_yr
        c_end   = C_FIRST + min((j+1)*months_per_yr, N) - 1
        st.total_cell(row, yr_start_col+j,
                      f"=SUM({_col(c_start)}{row}:{_col(c_end)}{row})",
                      nf, "grand")
    row += 1
    sep()

    # ── COST OF REVENUE ───────────────────────────────────────────────────────
    st.section_title(row, 1, "COST OF REVENUE", C_LAST + n_years + 1)
    ws.row_dimensions[row].height = 16
    row += 1

    R_COGS = row
    for i in range(N):
        col = C_FIRST + i
        st.formula_cell(row, col, f"=-{_col(col)}{R_TOT_REV}*{COGS_PCT}", nf)
    st.label(row, 1, "  Cost of Goods Sold / COGS", bold=False)
    for j in range(n_years):
        c_start = C_FIRST + j * months_per_yr
        c_end   = C_FIRST + min((j+1)*months_per_yr, N) - 1
        st.total_cell(row, yr_start_col+j,
                      f"=SUM({_col(c_start)}{row}:{_col(c_end)}{row})", nf, "sub")
    row += 1

    R_GROSS = row  # ← GROSS PROFIT
    for i in range(N):
        col = C_FIRST + i
        st.total_cell(row, col,
                      f"={_col(col)}{R_TOT_REV}+{_col(col)}{R_COGS}",
                      nf, "sub")
    st.label(row, 1, "GROSS PROFIT", bold=True)
    ws.cell(row=row, column=1).fill = _fill(TOTAL_BG)
    for j in range(n_years):
        c_start = C_FIRST + j * months_per_yr
        c_end   = C_FIRST + min((j+1)*months_per_yr, N) - 1
        st.total_cell(row, yr_start_col+j,
                      f"=SUM({_col(c_start)}{row}:{_col(c_end)}{row})", nf, "grand")
    row += 1

    R_GM_PCT = row  # Gross Margin %
    for i in range(N):
        col = C_FIRST + i
        rv  = _col(col)
        cell = ws.cell(row=row, column=col,
                       value=f"=IFERROR({rv}{R_GROSS}/{rv}{R_TOT_REV},0)")
        cell.fill = _fill(LIGHT_GR)
        cell.font = _font(size=9, color=DARK_GR, bold=True)
        cell.alignment = _align("right")
        cell.border = _border("thin", MID_GR)
        cell.number_format = pf
    st.label(row, 1, "  Gross Margin %", italic=True)
    for j in range(n_years):
        c_start = C_FIRST + j * months_per_yr
        c_end   = C_FIRST + min((j+1)*months_per_yr, N) - 1
        ann_r = ws.cell(row=row, column=yr_start_col+j)
        ann_r.value = (f"=IFERROR(SUM({_col(c_start)}{R_GROSS}:{_col(c_end)}{R_GROSS})/"
                       f"SUM({_col(c_start)}{R_TOT_REV}:{_col(c_end)}{R_TOT_REV}),0)")
        ann_r.fill = _fill(GOLD_LT); ann_r.font = _font(bold=True, size=9, color=NAVY)
        ann_r.alignment = _align("right"); ann_r.number_format = pf
    row += 1
    sep()

    # ── OPEX ──────────────────────────────────────────────────────────────────
    st.section_title(row, 1, "OPERATING EXPENSES", C_LAST + n_years + 1)
    ws.row_dimensions[row].height = 16
    row += 1

    opex_rows = {}
    opex_defs = [
        ("R_SAL",  "  Salaries & Benefits",      SAL_BASE, SAL_GR,  True),
        ("R_MKTG", "  Marketing & Advertising",   MKTG,     None,    False),
        ("R_TECH",  "  Technology / Cloud / SaaS", TECH,     None,    False),
        ("R_RENT",  "  Rent & Facilities",          RENT,     None,    False),
        ("R_GA",    "  General & Administrative",   GA,       None,    False),
        ("R_PROF",  "  Professional Fees",          PROF,     None,    False),
        ("R_OTH",   "  Other Operating Expenses",   OTHER_OPX,None,   False),
    ]

    for key, label, base_ref, growth_ref, has_growth in opex_defs:
        r_this = row
        opex_rows[key] = r_this
        for i in range(N):
            col = C_FIRST + i
            if has_growth and growth_ref:
                # Salary grows annually: every 12 months
                yr_idx = i // 12
                if yr_idx == 0:
                    formula = f"=-{base_ref}"
                else:
                    formula = (f"=-{base_ref}"
                               f"*(1+{growth_ref})^{yr_idx}")
            else:
                formula = f"=-{base_ref}"
            st.formula_cell(row, col, formula, nf)
        st.label(row, 1, label)
        for j in range(n_years):
            c_start = C_FIRST + j * months_per_yr
            c_end   = C_FIRST + min((j+1)*months_per_yr, N) - 1
            st.total_cell(row, yr_start_col+j,
                          f"=SUM({_col(c_start)}{row}:{_col(c_end)}{row})", nf, "sub")
        row += 1

    R_TOT_OPEX = row  # ← TOTAL OPEX
    for i in range(N):
        col = C_FIRST + i
        parts = "+".join([f"{_col(col)}{opex_rows[k]}" for k in opex_rows])
        st.total_cell(row, col, f"={parts}", nf, "sub")
    st.label(row, 1, "TOTAL OPEX", bold=True)
    ws.cell(row=row, column=1).fill = _fill(TOTAL_BG)
    for j in range(n_years):
        c_start = C_FIRST + j * months_per_yr
        c_end   = C_FIRST + min((j+1)*months_per_yr, N) - 1
        st.total_cell(row, yr_start_col+j,
                      f"=SUM({_col(c_start)}{row}:{_col(c_end)}{row})", nf, "grand")
    row += 1
    sep()

    # ── EBITDA ────────────────────────────────────────────────────────────────
    R_EBITDA = row
    for i in range(N):
        col = C_FIRST + i
        st.total_cell(row, col,
                      f"={_col(col)}{R_GROSS}+{_col(col)}{R_TOT_OPEX}",
                      nf, "grand")
    st.label(row, 1, "EBITDA", bold=True)
    ws.cell(row=row, column=1).fill = _fill(GRAND_BG)
    for j in range(n_years):
        c_start = C_FIRST + j * months_per_yr
        c_end   = C_FIRST + min((j+1)*months_per_yr, N) - 1
        st.total_cell(row, yr_start_col+j,
                      f"=SUM({_col(c_start)}{row}:{_col(c_end)}{row})", nf, "grand")
    row += 1

    R_EBITDA_PCT = row
    for i in range(N):
        col = C_FIRST + i
        cv  = _col(col)
        cell = ws.cell(row=row, column=col,
                       value=f"=IFERROR({cv}{R_EBITDA}/{cv}{R_TOT_REV},0)")
        cell.fill = _fill(LIGHT_GR); cell.font = _font(size=9, bold=True, color=DARK_GR)
        cell.alignment = _align("right"); cell.number_format = pf
        cell.border = _border("thin", MID_GR)
    st.label(row, 1, "  EBITDA Margin %", italic=True)
    for j in range(n_years):
        c_start = C_FIRST + j * months_per_yr
        c_end   = C_FIRST + min((j+1)*months_per_yr, N) - 1
        ann_r = ws.cell(row=row, column=yr_start_col+j)
        ann_r.value = (f"=IFERROR(SUM({_col(c_start)}{R_EBITDA}:{_col(c_end)}{R_EBITDA})/"
                       f"SUM({_col(c_start)}{R_TOT_REV}:{_col(c_end)}{R_TOT_REV}),0)")
        ann_r.fill = _fill(GOLD_LT); ann_r.font = _font(bold=True, size=9, color=NAVY)
        ann_r.alignment = _align("right"); ann_r.number_format = pf
    row += 1
    sep()

    # ── D&A + Interest + EBT + Tax + Net ─────────────────────────────────────
    st.section_title(row, 1, "BELOW THE LINE", C_LAST + n_years + 1)
    ws.row_dimensions[row].height = 16
    row += 1

    R_DEPR = row
    for i in range(N):
        col = C_FIRST + i
        st.formula_cell(row, col, f"=-{DEPR}", nf)
    st.label(row, 1, "  Depreciation & Amortisation")
    for j in range(n_years):
        c_start = C_FIRST + j * months_per_yr
        c_end   = C_FIRST + min((j+1)*months_per_yr, N) - 1
        st.total_cell(row, yr_start_col+j,
                      f"=SUM({_col(c_start)}{row}:{_col(c_end)}{row})", nf, "sub")
    row += 1

    R_INT = row
    for i in range(N):
        col = C_FIRST + i
        st.formula_cell(row, col, f"={INTEREST}", nf)
    st.label(row, 1, "  Interest Income / (Expense)")
    for j in range(n_years):
        c_start = C_FIRST + j * months_per_yr
        c_end   = C_FIRST + min((j+1)*months_per_yr, N) - 1
        st.total_cell(row, yr_start_col+j,
                      f"=SUM({_col(c_start)}{row}:{_col(c_end)}{row})", nf, "sub")
    row += 1

    R_EBT = row  # Earnings Before Tax
    for i in range(N):
        col = C_FIRST + i
        cv  = _col(col)
        st.total_cell(row, col,
                      f"={cv}{R_EBITDA}+{cv}{R_DEPR}+{cv}{R_INT}",
                      nf, "sub")
    st.label(row, 1, "EBT (Earnings Before Tax)", bold=True)
    ws.cell(row=row, column=1).fill = _fill(TOTAL_BG)
    for j in range(n_years):
        c_start = C_FIRST + j * months_per_yr
        c_end   = C_FIRST + min((j+1)*months_per_yr, N) - 1
        st.total_cell(row, yr_start_col+j,
                      f"=SUM({_col(c_start)}{row}:{_col(c_end)}{row})", nf, "sub")
    row += 1

    R_TAX = row
    for i in range(N):
        col = C_FIRST + i
        cv  = _col(col)
        # Tax only on positive EBT
        st.formula_cell(row, col,
                        f"=-MAX(0,{cv}{R_EBT})*{TAX_RATE}", nf)
    st.label(row, 1, "  Income Tax Provision")
    for j in range(n_years):
        c_start = C_FIRST + j * months_per_yr
        c_end   = C_FIRST + min((j+1)*months_per_yr, N) - 1
        st.total_cell(row, yr_start_col+j,
                      f"=SUM({_col(c_start)}{row}:{_col(c_end)}{row})", nf, "sub")
    row += 1

    R_NET = row  # ← NET PROFIT
    for i in range(N):
        col = C_FIRST + i
        cv  = _col(col)
        cell = ws.cell(row=row, column=col,
                       value=f"={cv}{R_EBT}+{cv}{R_TAX}")
        # Conditional: green if positive, red if negative
        cell.font      = _font(bold=True, size=9, color=NAVY)
        cell.alignment = _align("right")
        cell.border    = _top_bottom(GOLD, GOLD)
        cell.number_format = nf
        cell.fill = _fill(GREEN_LT)
    st.label(row, 1, "NET PROFIT / (LOSS)", bold=True)
    ws.cell(row=row, column=1).fill = _fill(GRAND_BG)
    ws.cell(row=row, column=1).font = _font(bold=True, size=10, color=NAVY)
    for j in range(n_years):
        c_start = C_FIRST + j * months_per_yr
        c_end   = C_FIRST + min((j+1)*months_per_yr, N) - 1
        st.total_cell(row, yr_start_col+j,
                      f"=SUM({_col(c_start)}{row}:{_col(c_end)}{row})", nf, "grand")
    row += 1

    R_NET_PCT = row
    for i in range(N):
        col = C_FIRST + i
        cv  = _col(col)
        cell = ws.cell(row=row, column=col,
                       value=f"=IFERROR({cv}{R_NET}/{cv}{R_TOT_REV},0)")
        cell.fill = _fill(LIGHT_GR); cell.font = _font(size=9, bold=True, color=DARK_GR)
        cell.alignment = _align("right"); cell.number_format = pf
        cell.border = _border("thin", MID_GR)
    st.label(row, 1, "  Net Profit Margin %", italic=True)
    for j in range(n_years):
        c_start = C_FIRST + j * months_per_yr
        c_end   = C_FIRST + min((j+1)*months_per_yr, N) - 1
        ann_r = ws.cell(row=row, column=yr_start_col+j)
        ann_r.value = (f"=IFERROR(SUM({_col(c_start)}{R_NET}:{_col(c_end)}{R_NET})/"
                       f"SUM({_col(c_start)}{R_TOT_REV}:{_col(c_end)}{R_TOT_REV}),0)")
        ann_r.fill = _fill(GOLD_LT); ann_r.font = _font(bold=True, size=9, color=NAVY)
        ann_r.alignment = _align("right"); ann_r.number_format = pf
    row += 1

    # ── Revenue vs EBITDA chart ────────────────────────────────────────────────
    try:
        chart = BarChart()
        chart.type    = "col"
        chart.grouping = "clustered"
        chart.title   = "Revenue vs EBITDA (Monthly)"
        chart.y_axis.title = f"Amount ({sym})"
        chart.x_axis.title = "Month"
        chart.style   = 10
        chart.width   = 18; chart.height = 11

        # Revenue data
        rev_ref = Reference(ws, min_col=C_FIRST, max_col=C_LAST,
                            min_row=R_TOT_REV, max_row=R_TOT_REV)
        rev_ser = BarChart()
        from openpyxl.chart import Series as CSer
        chart.add_data(Reference(ws, min_col=C_FIRST, max_col=C_LAST,
                                 min_row=R_TOT_REV, max_row=R_TOT_REV),
                       titles_from_data=False)
        chart.add_data(Reference(ws, min_col=C_FIRST, max_col=C_LAST,
                                 min_row=R_EBITDA, max_row=R_EBITDA),
                       titles_from_data=False)
        chart.series[0].title.v = "Revenue"
        chart.series[1].title.v = "EBITDA"

        cats = Reference(ws, min_col=C_FIRST, max_col=C_LAST,
                         min_row=3, max_row=3)
        chart.set_categories(cats)

        ws.add_chart(chart, f"A{row + 2}")
    except Exception:
        pass   # Chart optional — model still valid without it

    # Print setup
    ws.print_area = f"A1:{_col(C_LAST + n_years + 2)}{row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1

    return dict(
        R_TOT_REV=R_TOT_REV, R_GROSS=R_GROSS,
        R_EBITDA=R_EBITDA, R_NET=R_NET,
        C_FIRST=C_FIRST, C_LAST=C_LAST, N=N
    )

# ─── CASH FLOW SHEET ──────────────────────────────────────────────────────────

def _add_cashflow(wb: Workbook, params: Dict, periods: List[str],
                  pl_refs: Dict) -> None:
    ws = wb.create_sheet("💰 Cash Flow")
    ws.sheet_properties.tabColor = NAVY_MID
    st  = Styler(ws, params.get("currency_symbol", "₹"))
    sym = params.get("currency_symbol", "₹")
    nf  = _num_fmt(sym)
    N   = len(periods)
    ASM = "='📊 Assumptions'!"
    PL  = "='📈 P&L'!"

    C_FIRST = 2; C_LAST = N + 1
    ws.freeze_panes = "B4"
    st.set_col_widths({1: 34})
    for c in range(C_FIRST, C_LAST + 4):
        ws.column_dimensions[_col(c)].width = 13

    # Title
    ws.merge_cells(f"A1:{_col(C_LAST + 2)}1")
    c1 = ws.cell(row=1, column=1,
                 value=f"{params.get('company_name','Company')} — Cash Flow Statement")
    c1.fill = _fill(NAVY_MID); c1.font = _font(bold=True, size=12, color=WHITE)
    c1.alignment = _align("center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells(f"A2:{_col(C_LAST + 2)}2")
    c2 = ws.cell(row=2, column=1,
                 value="Indirect Method | Derives from P&L + Balance Sheet assumptions")
    c2.fill = _fill(SLATE); c2.font = _font(size=8.5, italic=True, color=GOLD_LT)
    c2.alignment = _align("center")

    # Period headers
    ws.cell(row=3, column=1, value="Cash Flow Item").fill = _fill(NAVY)
    ws.cell(row=3, column=1).font = _font(bold=True, size=9, color=WHITE)
    ws.cell(row=3, column=1).alignment = _align("center")
    for i, p in enumerate(periods):
        c = ws.cell(row=3, column=C_FIRST + i, value=p)
        c.fill = _fill(NAVY); c.font = _font(bold=True, size=8.5, color=WHITE)
        c.alignment = _align("center")
    ws.row_dimensions[3].height = 16

    row = 4
    pl_R = pl_refs  # aliases

    def sep():
        nonlocal row; ws.row_dimensions[row].height = 4; row += 1

    DEPR_REF  = f"{ASM}B26"
    AR_DAYS   = f"{ASM}B31"
    AP_DAYS   = f"{ASM}B32"
    CAPEX_REF = f"{ASM}B33"
    OPEN_CASH = f"{ASM}B30"

    # ── OPERATING ACTIVITIES ─────────────────────────────────────────────────
    st.section_title(row, 1, "A. CASH FROM OPERATING ACTIVITIES", C_LAST + 2)
    ws.row_dimensions[row].height = 16; row += 1

    R_NET_PROFIT = row
    for i in range(N):
        col = C_FIRST + i
        st.formula_cell(row, col,
                        f"={PL}{_col(col)}{pl_R['R_NET']}", nf)
    st.label(row, 1, "  Net Profit / (Loss)")
    row += 1

    R_ADD_BACK_DEPR = row
    for i in range(N):
        col = C_FIRST + i
        st.formula_cell(row, col, f"={DEPR_REF}", nf)
    st.label(row, 1, "  Add: Depreciation & Amortisation")
    row += 1

    R_AR_CHANGE = row  # Decrease/(Increase) in AR
    for i in range(N):
        col = C_FIRST + i
        if i == 0:
            formula = f"=-{PL}{_col(col)}{pl_R['R_TOT_REV']}*{AR_DAYS}/30"
        else:
            prev = _col(C_FIRST + i - 1)
            formula = (f"=-({PL}{_col(col)}{pl_R['R_TOT_REV']}"
                       f"-{PL}{prev}{pl_R['R_TOT_REV']})*{AR_DAYS}/30")
        st.formula_cell(row, col, formula, nf)
    st.label(row, 1, "  Change in Accounts Receivable")
    row += 1

    R_AP_CHANGE = row  # Increase/(Decrease) in AP
    for i in range(N):
        col = C_FIRST + i
        cogs_col = _col(col)
        if i == 0:
            formula = f"={PL}{cogs_col}{pl_R['R_GROSS']+1}*(-1)*{AP_DAYS}/30"
        else:
            prev_col = _col(C_FIRST + i - 1)
            # AP change = diff in COGS * payment days / 30
            formula = f"=({PL}{cogs_col}{pl_R['R_GROSS']+1}-{PL}{prev_col}{pl_R['R_GROSS']+1})*{AP_DAYS}/30"
        st.formula_cell(row, col, formula, nf)
    st.label(row, 1, "  Change in Accounts Payable")
    row += 1

    R_CFO = row  # ← OPERATING CASH FLOW
    for i in range(N):
        col = C_FIRST + i
        cv  = _col(col)
        st.total_cell(row, col,
                      f"={cv}{R_NET_PROFIT}+{cv}{R_ADD_BACK_DEPR}+{cv}{R_AR_CHANGE}+{cv}{R_AP_CHANGE}",
                      nf, "grand")
    st.label(row, 1, "NET OPERATING CASH FLOW", bold=True)
    ws.cell(row=row, column=1).fill = _fill(GRAND_BG)
    row += 1; sep()

    # ── INVESTING ACTIVITIES ──────────────────────────────────────────────────
    st.section_title(row, 1, "B. CASH FROM INVESTING ACTIVITIES", C_LAST + 2)
    ws.row_dimensions[row].height = 16; row += 1

    R_CAPEX = row
    for i in range(N):
        col = C_FIRST + i
        st.formula_cell(row, col, f"=-{CAPEX_REF}", nf)
    st.label(row, 1, "  Capital Expenditure (Capex)")
    row += 1

    R_CFI = row
    for i in range(N):
        col = C_FIRST + i
        st.total_cell(row, col, f"={_col(col)}{R_CAPEX}", nf, "sub")
    st.label(row, 1, "NET INVESTING CASH FLOW", bold=True)
    ws.cell(row=row, column=1).fill = _fill(TOTAL_BG)
    row += 1; sep()

    # ── FINANCING ACTIVITIES ──────────────────────────────────────────────────
    st.section_title(row, 1, "C. CASH FROM FINANCING ACTIVITIES", C_LAST + 2)
    ws.row_dimensions[row].height = 16; row += 1

    R_DEBT = row
    for i in range(N):
        col = C_FIRST + i
        st.formula_cell(row, col, "=0", nf)
    st.label(row, 1, "  Debt Repayment / (Drawdown)")
    row += 1

    R_EQUITY = row
    for i in range(N):
        col = C_FIRST + i
        st.formula_cell(row, col, "=0", nf)
    st.label(row, 1, "  Equity Raise / (Buyback)")
    row += 1

    R_CFF = row
    for i in range(N):
        col = C_FIRST + i
        st.total_cell(row, col,
                      f"={_col(col)}{R_DEBT}+{_col(col)}{R_EQUITY}",
                      nf, "sub")
    st.label(row, 1, "NET FINANCING CASH FLOW", bold=True)
    ws.cell(row=row, column=1).fill = _fill(TOTAL_BG)
    row += 1; sep()

    # ── NET CASH MOVEMENT ─────────────────────────────────────────────────────
    R_NET_CASH = row
    for i in range(N):
        col = C_FIRST + i
        cv  = _col(col)
        st.total_cell(row, col,
                      f"={cv}{R_CFO}+{cv}{R_CFI}+{cv}{R_CFF}",
                      nf, "grand")
    st.label(row, 1, "NET CASH MOVEMENT", bold=True)
    ws.cell(row=row, column=1).fill = _fill(GRAND_BG)
    row += 1

    # ── OPENING / CLOSING CASH ────────────────────────────────────────────────
    R_OPEN_CASH = row
    for i in range(N):
        col = C_FIRST + i
        if i == 0:
            formula = f"={OPEN_CASH}"
        else:
            formula = f"={_col(col - 1)}{R_OPEN_CASH + 1}"   # prev closing
        st.formula_cell(row, col, formula, nf, color=DARK_GR)
    st.label(row, 1, "  Opening Cash Balance")
    row += 1

    R_CLOSE_CASH = row
    for i in range(N):
        col = C_FIRST + i
        cv  = _col(col)
        cell = ws.cell(row=row, column=col,
                       value=f"={cv}{R_OPEN_CASH}+{cv}{R_NET_CASH}")
        cell.fill = _fill(GREEN_LT); cell.font = _font(bold=True, size=9, color=GREEN_POS)
        cell.alignment = _align("right"); cell.border = _top_bottom(GOLD, GOLD)
        cell.number_format = nf
    st.label(row, 1, "CLOSING CASH BALANCE", bold=True)
    ws.cell(row=row, column=1).fill = _fill(GREEN_LT)
    ws.cell(row=row, column=1).font = _font(bold=True, size=9, color=GREEN_POS)
    row += 1

    # Runway indicator
    R_RUNWAY = row
    for i in range(N):
        col = C_FIRST + i
        cv  = _col(col)
        # If closing cash < 0: "❌ Insolvent", else "✅ X months"
        # Simplified: just show closing cash formatted
        cell = ws.cell(row=row, column=col,
                       value=f'=IF({cv}{R_CLOSE_CASH}<0,"❌ Depleted","✅ Positive")')
        cell.fill = _fill(LIGHT_GR); cell.font = _font(size=8.5, color=DARK_GR)
        cell.alignment = _align("center")
        cell.border = _border("thin", MID_GR)
    st.label(row, 1, "  Cash Status")

    # Chart
    try:
        lc = LineChart()
        lc.title  = "Closing Cash Balance Trend"
        lc.y_axis.title = f"Cash ({sym})"
        lc.x_axis.title = "Month"
        lc.style  = 10
        lc.width  = 18; lc.height = 11

        lc.add_data(Reference(ws, min_col=C_FIRST, max_col=C_LAST,
                               min_row=R_CLOSE_CASH, max_row=R_CLOSE_CASH),
                    titles_from_data=False)
        lc.series[0].title.v = "Closing Cash"
        lc.set_categories(Reference(ws, min_col=C_FIRST, max_col=C_LAST,
                                     min_row=3, max_row=3))
        ws.add_chart(lc, f"A{row + 3}")
    except Exception:
        pass

    ws.print_area = f"A1:{_col(C_LAST + 2)}{row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage  = True


# ─── DASHBOARD SHEET ──────────────────────────────────────────────────────────

def _add_dashboard(wb: Workbook, params: Dict, periods: List[str],
                   pl_refs: Dict) -> None:
    ws  = wb.create_sheet("📊 Dashboard")
    ws.sheet_properties.tabColor = "1A7A4A"
    st  = Styler(ws, params.get("currency_symbol", "₹"))
    sym = params.get("currency_symbol", "₹")
    nf  = _num_fmt(sym)
    pf  = _pct_fmt()
    N   = len(periods)
    PL  = "='📈 P&L'!"
    CF  = "='💰 Cash Flow'!"
    ASM = "='📊 Assumptions'!"
    C_FIRST = pl_refs["C_FIRST"]
    C_LAST  = pl_refs["C_LAST"]
    R = pl_refs

    ws.freeze_panes = "A3"
    st.set_col_widths({1: 26, 2: 20, 3: 20, 4: 20, 5: 20,
                       6: 20, 7: 20, 8: 20})

    # Big title banner
    ws.merge_cells("A1:H1")
    c1 = ws.cell(row=1, column=1,
                 value=f"  {params.get('company_name','Company')} — Executive Financial Dashboard")
    c1.fill = _fill(NAVY); c1.font = _font(bold=True, size=14, color=WHITE)
    c1.alignment = _align("left")
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:H2")
    date_gen = datetime.now().strftime("%d %b %Y %H:%M")
    c2 = ws.cell(row=2, column=1,
                 value=f"  Auto-calculated from P&L & Cash Flow  ·  Generated: {date_gen}  ·  Currency: {params.get('currency','INR')}")
    c2.fill = _fill(SLATE); c2.font = _font(size=8.5, italic=True, color=GOLD_LT)
    ws.row_dimensions[2].height = 16

    # ── KPI Cards (row 4–9) ───────────────────────────────────────────────────
    # Derive columns for Year 1 and Year 2 range
    Y1_start = C_FIRST; Y1_end = C_FIRST + min(12, N) - 1
    Y2_start = C_FIRST + 12; Y2_end = C_FIRST + min(24, N) - 1

    def kpi_block(row, col, label, value_formula, prior_formula=None,
                  fmt=None, bg=NAVY, val_color=WHITE, label_color=GOLD_LT):
        """Draw KPI card block (label row + value row + change row)."""
        # Header
        lbl = ws.cell(row=row, column=col, value=label)
        lbl.fill = _fill(bg); lbl.font = _font(bold=True, size=7.5, color=label_color)
        lbl.alignment = _align("center")
        ws.row_dimensions[row].height = 13

        # Value
        val = ws.cell(row=row+1, column=col, value=value_formula)
        val.fill = _fill(bg); val.font = _font(bold=True, size=14, color=val_color)
        val.alignment = _align("center")
        if fmt: val.number_format = fmt
        ws.row_dimensions[row+1].height = 22

        # Change vs prior
        if prior_formula:
            chg = ws.cell(row=row+2, column=col,
                          value=f'=IFERROR(({value_formula})/({prior_formula})-1,"—")')
            chg.fill = _fill(bg); chg.font = _font(size=8, color=GOLD_LT)
            chg.alignment = _align("center"); chg.number_format = pf
        else:
            ws.cell(row=row+2, column=col, value="YTD").fill = _fill(bg)
        ws.row_dimensions[row+2].height = 14

    row_kpi = 4
    # 6 KPI cards across cols 2-7
    half = min(12, N)
    kpi_defs = [
        ("📈 Revenue (YTD)",  f"=SUM({PL}{_col(Y1_start)}{R['R_TOT_REV']}:{PL}{_col(Y1_end)}{R['R_TOT_REV']})", None, nf),
        ("💰 Gross Profit",   f"=SUM({PL}{_col(Y1_start)}{R['R_GROSS']}:{PL}{_col(Y1_end)}{R['R_GROSS']})",     None, nf),
        ("📊 EBITDA",         f"=SUM({PL}{_col(Y1_start)}{R['R_EBITDA']}:{PL}{_col(Y1_end)}{R['R_EBITDA']})",   None, nf),
        ("💵 Net Profit",     f"=SUM({PL}{_col(Y1_start)}{R['R_NET']}:{PL}{_col(Y1_end)}{R['R_NET']})",         None, nf),
        ("% GM",              f"=IFERROR(SUM({PL}{_col(Y1_start)}{R['R_GROSS']}:{PL}{_col(Y1_end)}{R['R_GROSS']})/SUM({PL}{_col(Y1_start)}{R['R_TOT_REV']}:{PL}{_col(Y1_end)}{R['R_TOT_REV']}),0)", None, pf),
        ("🏦 Cash (EOM)",     f"={CF}{_col(C_LAST)}{8}",  None, nf),
    ]
    for j, (label, val_f, prior_f, fmt) in enumerate(kpi_defs):
        kpi_block(row_kpi, j+2, label, val_f, prior_f, fmt,
                  bg=NAVY if j < 4 else NAVY_MID)

    # Divider
    row_after_kpi = row_kpi + 4
    ws.merge_cells(f"A{row_after_kpi}:H{row_after_kpi}")
    div = ws.cell(row=row_after_kpi, column=1)
    div.fill = _fill(GOLD); ws.row_dimensions[row_after_kpi].height = 3

    # ── Annual Summary Table ──────────────────────────────────────────────────
    row_tbl = row_after_kpi + 2
    n_years = len(_year_labels(periods))
    yr_labels = _year_labels(periods)
    months_per_yr = N // max(1, n_years)

    ws.cell(row=row_tbl, column=1, value="ANNUAL SUMMARY").fill = _fill(NAVY_MID)
    ws.cell(row=row_tbl, column=1).font = _font(bold=True, size=9, color=WHITE)
    ws.cell(row=row_tbl, column=1).alignment = _align("left")
    hdr_cols = [1] + list(range(2, 2 + n_years + 2))  # label + years + CAGR + Last Month
    for ci, hdr in enumerate(["Metric"] + yr_labels + ["CAGR", "Latest Month"]):
        c = ws.cell(row=row_tbl, column=ci + 1, value=hdr)
        c.fill = _fill(NAVY); c.font = _font(bold=True, size=8.5, color=WHITE)
        c.alignment = _align("center")
    ws.row_dimensions[row_tbl].height = 16

    summary_metrics = [
        ("Revenue",       R["R_TOT_REV"], nf),
        ("Gross Profit",  R["R_GROSS"],   nf),
        ("GM %",          None,           pf),  # special
        ("EBITDA",        R["R_EBITDA"],  nf),
        ("EBITDA %",      None,           pf),
        ("Net Profit",    R["R_NET"],     nf),
        ("Net Margin %",  None,           pf),
    ]

    row_tbl += 1
    is_alt = False
    for metric_name, pl_row, fmt in summary_metrics:
        bg = LIGHT_GR if is_alt else WHITE
        ws.cell(row=row_tbl, column=1, value=metric_name).fill = _fill(TOTAL_BG)
        ws.cell(row=row_tbl, column=1).font = _font(bold=True, size=9, color=NAVY)
        ws.cell(row=row_tbl, column=1).alignment = _align("left")
        ws.row_dimensions[row_tbl].height = 15

        for j in range(n_years):
            c_start = C_FIRST + j * months_per_yr
            c_end   = C_FIRST + min((j+1)*months_per_yr, N) - 1
            col = j + 2

            if metric_name in ("GM %", "EBITDA %", "Net Margin %"):
                num_row = R["R_GROSS"] if "GM" in metric_name else (
                    R["R_EBITDA"] if "EBITDA" in metric_name else R["R_NET"])
                fml = (f"=IFERROR(SUM({PL}{_col(c_start)}{num_row}:{PL}{_col(c_end)}{num_row})/"
                       f"SUM({PL}{_col(c_start)}{R['R_TOT_REV']}:{PL}{_col(c_end)}{R['R_TOT_REV']}),0)")
            else:
                fml = f"=SUM({PL}{_col(c_start)}{pl_row}:{PL}{_col(c_end)}{pl_row})"

            cell = ws.cell(row=row_tbl, column=col, value=fml)
            cell.fill = _fill(bg); cell.font = _font(size=9, color=DARK_GR)
            cell.alignment = _align("right"); cell.number_format = fmt
            cell.border = _border("thin", MID_GR)

        # CAGR (if ≥2 years)
        if n_years >= 2 and pl_row and metric_name not in ("GM %","EBITDA %","Net Margin %"):
            yr1_c = 2; yr_last_c = 1 + n_years
            cagr_cell = ws.cell(row=row_tbl, column=2 + n_years)
            cagr_cell.value = (f"=IFERROR((ABS({_col(yr_last_c+1)}{row_tbl}/"
                               f"{_col(yr1_c)}{row_tbl}))^(1/{n_years-1})-1,\"—\")")
            cagr_cell.fill = _fill(AMBER_LT); cagr_cell.font = _font(bold=True, size=9, color=AMBER)
            cagr_cell.alignment = _align("center"); cagr_cell.number_format = pf
        else:
            ws.cell(row=row_tbl, column=2+n_years, value="—")

        # Latest month
        if pl_row and metric_name not in ("GM %","EBITDA %","Net Margin %"):
            lm = ws.cell(row=row_tbl, column=3+n_years,
                         value=f"={PL}{_col(C_LAST)}{pl_row}")
            lm.fill = _fill(bg); lm.font = _font(size=9)
            lm.alignment = _align("right"); lm.number_format = fmt

        row_tbl += 1
        is_alt = not is_alt

    ws.print_area = f"A1:H{row_tbl + 2}"
    ws.page_setup.fitToPage  = True
    ws.page_setup.fitToWidth = 1

# ─── MAIN DISPATCH ────────────────────────────────────────────────────────────

def build_excel(params: Dict[str, Any]) -> bytes:
    """Entry point. Returns .xlsx bytes."""
    A   = params.get("assumptions", {})
    N   = int(_num(A.get("num_periods", params.get("num_periods", 12))))
    N   = max(3, min(N, 36))
    sm  = A.get("start_month", params.get("start_month", "Apr-25"))
    periods = _periods(sm, N)

    wb = Workbook()
    wb.remove(wb.active)   # remove default blank sheet

    _add_instructions(wb, params)
    _add_assumptions(wb, params, periods)
    pl_refs = _add_pl(wb, params, periods)
    _add_cashflow(wb, params, periods, pl_refs)
    _add_dashboard(wb, params, periods, pl_refs)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
