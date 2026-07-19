"""
OrchestrIQ Excel Engine v3 — CFO/Board Grade
Uses openpyxl for real styling, formulas, charts, freeze panes, conditional formatting.
AI extracts structured data. Python builds the workbook deterministically.
"""
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side,
    GradientFill, numbers as xl_numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
import io, re
from datetime import datetime, timedelta

# ── Brand Palette ──────────────────────────────────────────────────────────────
NAVY    = "1E3A5F"
TEAL    = "14B8A6"
LIGHT   = "F1F5F9"
WHITE   = "FFFFFF"
MUTED   = "94A3B8"
DARK    = "0F172A"
GREEN   = "10B981"
AMBER   = "F59E0B"
RED     = "EF4444"
ROW_ALT = "F8FAFC"
BORDER  = "E2E8F0"

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, color=DARK, size=11, italic=False, name="Calibri"):
    return Font(bold=bold, color=color, size=size, italic=italic, name=name)

def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _border(color=BORDER):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def _header_cell(ws, row, col, value, bg=NAVY, fg=WHITE, size=11, bold=True, align="left"):
    c = ws.cell(row=row, column=col, value=value)
    c.fill = _fill(bg)
    c.font = _font(bold=bold, color=fg, size=size)
    c.alignment = _align(align)
    c.border = _border(bg)
    return c

def _data_cell(ws, row, col, value, fmt=None, bold=False, color=DARK, bg=WHITE, align="left"):
    c = ws.cell(row=row, column=col, value=value)
    c.font = _font(bold=bold, color=color)
    c.alignment = _align(align)
    c.border = _border()
    if row % 2 == 0:
        c.fill = _fill(ROW_ALT)
    if fmt:
        c.number_format = fmt
    return c

def _autofit_columns(ws, min_w=10, max_w=50):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                v = str(cell.value) if cell.value is not None else ""
                max_len = max(max_len, len(v))
            except:
                pass
        ws.column_dimensions[col_letter].width = min(max_w, max(min_w, max_len + 3))

def _num(s):
    """Parse a value to float if possible."""
    if isinstance(s, (int, float)):
        return float(s)
    try:
        return float(str(s).replace(",", "").replace("₹", "").replace("$", "").replace("%", "").strip())
    except:
        return None

def build_excel(schema: dict, currency_symbol: str = "₹") -> bytes:
    """
    Main entry. schema must have:
      title, company, sheets: [{name, type, headers, rows, summary_kpis}]
    """
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    sheets_data = schema.get("sheets", [])
    if not sheets_data:
        raise ValueError("No sheets in schema")

    # ── 1. COVER / DASHBOARD ──────────────────────────────────────────────────
    cover = wb.create_sheet("📊 Dashboard", 0)
    cover.sheet_view.showGridLines = False
    cover.row_dimensions[1].height = 60
    cover.row_dimensions[2].height = 30

    # Title band
    cover.merge_cells("A1:H1")
    t = cover["A1"]
    t.value = schema.get("title", "Executive Report")
    t.fill = _fill(NAVY)
    t.font = Font(name="Calibri", bold=True, size=22, color=WHITE)
    t.alignment = _align("center", "center")

    # Sub-line
    cover.merge_cells("A2:H2")
    s = cover["A2"]
    s.value = f"{schema.get('company','')or''}  ·  {schema.get('industry','')or''}  ·  Generated {datetime.now().strftime('%d %b %Y')}"
    s.fill = _fill(TEAL)
    s.font = Font(name="Calibri", size=11, color=WHITE, italic=True)
    s.alignment = _align("center", "center")

    cover.row_dimensions[3].height = 8

    # ── KPI Cards row ─────────────────────────────────────────────────────────
    kpis = schema.get("summary_kpis", [])
    if not kpis:
        # Auto-extract from first data sheet
        for sh in sheets_data:
            if sh.get("summary_kpis"):
                kpis = sh["summary_kpis"]
                break
    
    kpi_row = 4
    col = 1
    for kpi in kpis[:6]:
        label = kpi.get("label","")
        value = kpi.get("value","")
        delta = kpi.get("delta","")
        
        cover.merge_cells(start_row=kpi_row, start_column=col, end_row=kpi_row, end_column=col+1)
        cover.merge_cells(start_row=kpi_row+1, start_column=col, end_row=kpi_row+1, end_column=col+1)
        cover.merge_cells(start_row=kpi_row+2, start_column=col, end_row=kpi_row+2, end_column=col+1)
        cover.merge_cells(start_row=kpi_row+3, start_column=col, end_row=kpi_row+3, end_column=col+1)

        # Border box
        for r2 in range(kpi_row, kpi_row+4):
            for c2 in range(col, col+2):
                cell = cover.cell(row=r2, column=c2)
                cell.fill = _fill(LIGHT)
                cell.border = Border(
                    left=Side(style="medium",color=TEAL) if c2==col else Side(style="thin",color=BORDER),
                    right=Side(style="thin",color=BORDER),
                    top=Side(style="thin",color=BORDER) if r2==kpi_row else Side(style="thin",color=BORDER),
                    bottom=Side(style="medium",color=BORDER) if r2==kpi_row+3 else Side(style="thin",color=BORDER),
                )
        
        lc = cover.cell(row=kpi_row, column=col, value=label)
        lc.font = Font(name="Calibri", size=9, color=MUTED, bold=True)
        lc.alignment = _align("left","bottom")

        vc = cover.cell(row=kpi_row+1, column=col, value=value)
        vc.font = Font(name="Calibri", size=18, bold=True, color=NAVY)
        vc.alignment = _align("left","center")

        if delta:
            dc = cover.cell(row=kpi_row+2, column=col, value=str(delta))
            is_pos = "+" in str(delta) or (str(delta).startswith("-") is False and str(delta) != "0")
            dc.font = Font(name="Calibri", size=10, bold=True,
                          color=GREEN if "+" in str(delta) else RED if "-" in str(delta) else MUTED)
            dc.alignment = _align("left","center")

        col += 2
        if col > 12:
            break

    # ── 2. DATA SHEETS ────────────────────────────────────────────────────────
    chart_sheets = []
    for sheet_spec in sheets_data:
        sname = str(sheet_spec.get("name","Sheet"))[:31]
        stype = sheet_spec.get("type","data")
        ws = wb.create_sheet(sname)
        ws.sheet_view.showGridLines = True

        headers = sheet_spec.get("headers", [])
        rows = sheet_spec.get("rows", [])
        
        if not headers and rows:
            headers = [f"Column {i+1}" for i in range(len(rows[0]) if rows else 0)]

        # Write headers
        for ci, h in enumerate(headers, 1):
            _header_cell(ws, 1, ci, h)

        # Freeze header row
        ws.freeze_panes = "A2"

        # Write data rows
        numeric_cols = set()
        pct_cols = set()
        for ri, row in enumerate(rows, 2):
            for ci, val in enumerate(row, 1):
                # Detect percentage columns by header name
                h = headers[ci-1] if ci <= len(headers) else ""
                is_pct = any(k in h.lower() for k in ["%","percent","margin","rate","growth","ratio"])
                
                n = _num(val)
                if n is not None and not isinstance(val, str):
                    numeric_cols.add(ci)
                    if is_pct:
                        pct_cols.add(ci)
                    
                    fmt = "0.00%" if is_pct else f'"{currency_symbol}"#,##0.00' if ci in numeric_cols else "General"
                    _data_cell(ws, ri, ci, n if not is_pct else n/100 if n > 1 else n, fmt=fmt)
                elif isinstance(val, str) and val.strip().startswith("="):
                    c = ws.cell(row=ri, column=ci, value=None)
                    c.value = val.strip()
                    c.font = _font()
                    c.alignment = _align()
                    c.border = _border()
                    if ri % 2 == 0:
                        c.fill = _fill(ROW_ALT)
                    numeric_cols.add(ci)
                else:
                    _data_cell(ws, ri, ci, val)

        # Add totals row
        if rows and len(rows) > 1:
            total_row = len(rows) + 2
            ws.cell(row=total_row, column=1, value="TOTAL").font = _font(bold=True, color=WHITE)
            ws.cell(row=total_row, column=1).fill = _fill(NAVY)
            ws.cell(row=total_row, column=1).alignment = _align()
            
            for ci in numeric_cols:
                if ci > len(headers):
                    continue
                col_letter = get_column_letter(ci)
                formula = f"=SUM({col_letter}2:{col_letter}{total_row-1})"
                c = ws.cell(row=total_row, column=ci, value=formula)
                c.font = _font(bold=True, color=WHITE)
                c.fill = _fill(NAVY)
                c.alignment = _align("right")
                h = headers[ci-1] if ci <= len(headers) else ""
                is_pct = any(k in h.lower() for k in ["%","percent","margin","rate","growth"])
                c.number_format = "0.00%" if is_pct else f'"{currency_symbol}"#,##0.00'

        # Conditional formatting on variance/delta columns
        for ci, h in enumerate(headers, 1):
            if any(k in h.lower() for k in ["variance","delta","change","growth","diff"]) and rows:
                col_letter = get_column_letter(ci)
                last_row = len(rows) + 1
                range_str = f"{col_letter}2:{col_letter}{last_row}"
                ws.conditional_formatting.add(range_str,
                    ColorScaleRule(
                        start_type="min", start_color=RED,
                        mid_type="num", mid_value=0, mid_color=AMBER,
                        end_type="max", end_color=GREEN
                    )
                )

        # Auto-filter
        if headers and rows:
            ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows)+1}"

        # Column widths
        _autofit_columns(ws)

        # Register for chart generation
        if len(rows) >= 3 and len(numeric_cols) >= 1:
            chart_sheets.append((ws, sname, headers, rows, numeric_cols))

    # ── 3. CHARTS SHEET ───────────────────────────────────────────────────────
    if chart_sheets:
        try:
            cws = wb.create_sheet("📈 Charts")
            cws.sheet_view.showGridLines = False
            _header_cell(cws, 1, 1, "VISUAL DASHBOARD", bg=NAVY, fg=WHITE, size=14)
            cws.merge_cells("A1:L1")
            cws.row_dimensions[1].height = 30

            chart_row = 3
            for (src_ws, sname, headers, rows, numeric_cols) in chart_sheets[:3]:
                # Find best label col (col 1) and first numeric col
                num_col = sorted(numeric_cols)[0] if numeric_cols else 2
                last_data_row = len(rows) + 1

                # Bar chart
                chart = BarChart()
                chart.type = "col"
                chart.grouping = "clustered"
                chart.title = sname
                chart.style = 10
                chart.y_axis.title = headers[num_col-1] if num_col <= len(headers) else "Value"
                chart.x_axis.title = headers[0]
                chart.width = 20
                chart.height = 12

                data_ref = Reference(src_ws, min_col=num_col, min_row=1, max_row=last_data_row)
                cats_ref = Reference(src_ws, min_col=1, min_row=2, max_row=last_data_row)
                chart.add_data(data_ref, titles_from_data=True)
                chart.set_categories(cats_ref)
                chart.series[0].graphicalProperties.solidFill = TEAL

                cws.add_chart(chart, f"A{chart_row}")
                chart_row += 22

                # Line chart if time-series pattern
                if any(any(k in str(r[0]).lower() for k in ["jan","feb","q1","q2","2024","2025","2026"]) for r in rows[:3]):
                    line = LineChart()
                    line.title = f"{sname} — Trend"
                    line.style = 10
                    line.width = 20
                    line.height = 12
                    line.add_data(data_ref, titles_from_data=True)
                    line.set_categories(cats_ref)
                    line.series[0].graphicalProperties.line.solidFill = NAVY
                    cws.add_chart(line, f"L{chart_row - 22}")

        except Exception as e:
            pass  # Charts are additive — never block delivery

    # ── 4. ASSUMPTIONS SHEET ─────────────────────────────────────────────────
    assumptions = schema.get("assumptions", [])
    if assumptions:
        aws = wb.create_sheet("⚙ Assumptions")
        aws.sheet_view.showGridLines = True
        _header_cell(aws, 1, 1, "Parameter")
        _header_cell(aws, 1, 2, "Value")
        _header_cell(aws, 1, 3, "Basis")
        _header_cell(aws, 1, 4, "Confidence")
        aws.freeze_panes = "A2"
        for ri, a in enumerate(assumptions, 2):
            if isinstance(a, dict):
                _data_cell(aws, ri, 1, a.get("parameter",""))
                _data_cell(aws, ri, 2, a.get("value",""))
                _data_cell(aws, ri, 3, a.get("basis",""))
                _data_cell(aws, ri, 4, a.get("confidence","[ESTIMATE]"))
            else:
                _data_cell(aws, ri, 1, str(a))
        _autofit_columns(aws)

    # ── 5. INSTRUCTIONS SHEET ─────────────────────────────────────────────────
    iws = wb.create_sheet("ℹ Instructions")
    iws.sheet_view.showGridLines = False
    iws.merge_cells("A1:D1")
    _header_cell(iws, 1, 1, "HOW TO USE THIS WORKBOOK", bg=NAVY, fg=WHITE, size=13)
    instructions = schema.get("instructions","")
    lines = [
        f"Generated: {datetime.now().strftime('%d %B %Y, %H:%M')}",
        f"Currency: {currency_symbol}",
        "",
        "NAVIGATION:",
        "  📊 Dashboard   — Key metrics overview",
        "  📈 Charts      — Visual trend analysis",
        "  ⚙ Assumptions  — Edit input variables here (other sheets update automatically)",
        "  ℹ Instructions — This sheet",
        "",
        "USAGE RULES:",
        "  • Green cells = input cells (edit these)",
        "  • Blue header = formula cells (do not edit)",
        "  • Red/Amber/Green colors = variance heat map",
        "  • All monetary values in " + currency_symbol,
        "",
    ]
    if instructions:
        lines.append("SPECIFIC NOTES:")
        for line in str(instructions).split("\n"):
            if line.strip():
                lines.append("  " + line.strip())
    
    for ri, line in enumerate(lines, 2):
        c = iws.cell(row=ri, column=1, value=line)
        c.font = Font(name="Calibri", size=11, color=DARK,
                     bold=line.endswith(":") or line.startswith("Generated"))
        c.alignment = _align()
    _autofit_columns(iws)

    # Output
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
