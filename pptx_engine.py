"""
OrchestrIQ PPTX Engine — CEO/Board-Grade Presentation Builder
python-pptx: 13.33" × 7.5" widescreen, all slides fully custom (blank layout).
"""

import io
from datetime import datetime
from typing import Any, Dict, List, Optional

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt

# ─── BRAND PALETTE ────────────────────────────────────────────────────────────
def _rgb(r, g, b): return RGBColor(r, g, b)

NAVY     = _rgb(13,  27,  42)
NAVY_MID = _rgb(26,  39,  68)
SLATE    = _rgb(44,  62,  80)
GOLD     = _rgb(201, 168, 76)
GOLD_LT  = _rgb(240, 208, 128)
WHITE    = _rgb(255, 255, 255)
LIGHT_GR = _rgb(247, 248, 250)
MID_GR   = _rgb(214, 216, 219)
DARK_GR  = _rgb(58,  58,  58)
GREEN_OK = _rgb(26,  122, 74)
GREEN_LT = _rgb(230, 244, 237)
RED_NG   = _rgb(192, 57,  43)
RED_LT   = _rgb(253, 232, 230)
AMBER    = _rgb(214, 137, 16)
BLUE_AC  = _rgb(31,  78,  121)
TEAL_AC  = _rgb(13,  110, 138)

# ─── SLIDE DIMENSIONS ─────────────────────────────────────────────────────────
SW = Inches(13.33)   # Slide Width
SH = Inches(7.5)     # Slide Height

# ─── HELPERS ──────────────────────────────────────────────────────────────────

def _num(v):
    try: return float(str(v or 0).replace(",","").replace("%",""))
    except: return 0.0

def _fmt(v, sym="₹"):
    v = _num(v)
    if abs(v) >= 1e7:  return f"{sym}{v/1e7:.1f}Cr"
    if abs(v) >= 1e5:  return f"{sym}{v/1e5:.1f}L"
    if abs(v) >= 1000: return f"{sym}{v/1000:.1f}K"
    return f"{sym}{v:.0f}"

def _fmt_pct(v):
    return f"{_num(v):.1f}%"

def _blank_slide(prs: Presentation) -> Any:
    """Add a new slide using blank layout."""
    layout = prs.slide_layouts[6]   # index 6 = Blank in most themes
    return prs.slides.add_slide(layout)

def _rect(slide, l, t, w, h, fill_color: RGBColor = None,
          line_color: RGBColor = None, line_width: int = 0):
    """Add a filled/stroked rectangle shape."""
    shape = slide.shapes.add_shape(
        1, l, t, w, h)   # MSO_SHAPE_TYPE.RECTANGLE = 1
    shape.fill.solid() if fill_color else shape.fill.background()
    if fill_color:
        shape.fill.fore_color.rgb = fill_color
    if line_color and line_width:
        shape.line.color.rgb = line_color
        shape.line.width     = Pt(line_width)
    else:
        shape.line.fill.background()
    return shape

def _text_box(slide, text, l, t, w, h,
              size=12, bold=False, color=WHITE,
              align=PP_ALIGN.LEFT, wrap=True, italic=False,
              name="Calibri"):
    """Add text box."""
    txb = slide.shapes.add_textbox(l, t, w, h)
    tf  = txb.text_frame
    tf.word_wrap = wrap
    tf.auto_size = None
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text)
    run.font.name   = name
    run.font.size   = Pt(size)
    run.font.bold   = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txb

def _add_para(tf, text, size=10, bold=False, color=WHITE,
              align=PP_ALIGN.LEFT, space_before=Pt(4), indent=False):
    """Add paragraph to existing text frame."""
    p = tf.add_paragraph()
    p.alignment = align
    if space_before:
        p.space_before = space_before
    if indent:
        p.level = 1
    run = p.add_run()
    run.text = str(text)
    run.font.size  = Pt(size)
    run.font.bold  = bold
    run.font.color.rgb = color
    return p

# ─── GLOBAL HEADER / FOOTER DECORATION ───────────────────────────────────────

def _add_slide_chrome(slide, company: str, classification: str,
                      slide_num: int, total: int, is_title=False):
    """Add header bar (company name + classification) and footer (slide #)."""
    if is_title:
        # Title slides don't need the header chrome
        # But add classification + date at bottom
        _text_box(slide, f"{classification}  ·  {datetime.now().strftime('%B %Y')}",
                  Inches(0.3), SH - Inches(0.4), SW - Inches(0.6), Inches(0.3),
                  size=7, color=MID_GR, align=PP_ALIGN.CENTER)
        return

    # Header bar — thin navy strip at very top
    _rect(slide, 0, 0, SW, Inches(0.38), NAVY)
    # Gold left accent
    _rect(slide, 0, 0, Inches(0.07), Inches(0.38), GOLD)
    # Company name (left)
    _text_box(slide, company.upper(), Inches(0.15), Inches(0.04),
              Inches(6), Inches(0.3), size=7, bold=True, color=WHITE)
    # Classification (right)
    _text_box(slide, classification, SW - Inches(2.5), Inches(0.04),
              Inches(2.3), Inches(0.3), size=7, color=GOLD_LT,
              align=PP_ALIGN.RIGHT)

    # Gold bottom line
    _rect(slide, 0, SH - Inches(0.06), SW, Inches(0.06), GOLD)
    # Footer: slide number
    _text_box(slide, f"{slide_num} / {total}",
              SW - Inches(1.0), SH - Inches(0.36), Inches(0.85), Inches(0.28),
              size=7.5, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
    # Footer: date
    _text_box(slide, datetime.now().strftime("%B %Y"),
              Inches(0.2), SH - Inches(0.36), Inches(2), Inches(0.28),
              size=7.5, color=SLATE)

# ─── SLIDE BUILDERS ───────────────────────────────────────────────────────────

def _slide_title(prs, params):
    """Full navy title slide."""
    slide = _blank_slide(prs)
    company      = params.get("company_name", "Company")
    title        = params.get("title",        "Presentation Title")
    subtitle     = params.get("subtitle",     "")
    classification = params.get("classification", "CONFIDENTIAL")
    date_str     = params.get("date", datetime.now().strftime("%B %Y"))
    audience     = params.get("audience",     "Executive Management")

    # Full navy bg
    _rect(slide, 0, 0, SW, SH, NAVY)
    # Gold left stripe
    _rect(slide, 0, 0, Inches(0.18), SH, GOLD)
    # Dark bottom panel
    _rect(slide, 0, SH - Inches(2.2), SW, Inches(2.2), NAVY_MID)
    # Gold divider above dark panel
    _rect(slide, 0, SH - Inches(2.25), SW, Inches(0.04), GOLD)

    # Company name
    _text_box(slide, company.upper(), Inches(0.38), Inches(0.5),
              SW - Inches(0.6), Inches(0.5),
              size=10, bold=True, color=GOLD_LT)

    # Classification badge
    _rect(slide, SW - Inches(2.0), Inches(0.4), Inches(1.7), Inches(0.36), GOLD)
    _text_box(slide, classification, SW - Inches(2.0), Inches(0.4),
              Inches(1.7), Inches(0.36), size=8, bold=True, color=NAVY,
              align=PP_ALIGN.CENTER)

    # Main title (large, white)
    _text_box(slide, title, Inches(0.38), Inches(1.5),
              SW - Inches(0.6), Inches(2.8),
              size=32, bold=True, color=WHITE)

    # Gold divider line drawn via shape
    _rect(slide, Inches(0.38), Inches(4.5), SW - Inches(0.6), Inches(0.03), GOLD)

    # Subtitle
    if subtitle:
        _text_box(slide, subtitle, Inches(0.38), Inches(4.65),
                  SW - Inches(0.6), Inches(0.65),
                  size=14, italic=True, color=GOLD_LT)

    # Bottom panel: prepared for / date
    _text_box(slide, f"Prepared for: {audience}",
              Inches(0.38), SH - Inches(1.9), Inches(6), Inches(0.4),
              size=9.5, color=MID_GR)
    _text_box(slide, date_str,
              Inches(0.38), SH - Inches(1.5), Inches(4), Inches(0.4),
              size=9.5, bold=True, color=GOLD_LT)

    # OrchestrIQ branding
    _text_box(slide, "Generated by OrchestrIQ | GorakhAI",
              SW - Inches(3.2), SH - Inches(0.5),
              Inches(3.0), Inches(0.35), size=7, color=SLATE,
              align=PP_ALIGN.RIGHT)

    _add_slide_chrome(slide, company, classification, 1, 1, is_title=True)


def _slide_exec_summary(prs, slide_data, params, num, total):
    """3-column insight card slide."""
    slide = _blank_slide(prs)
    company = params.get("company_name", "")
    sym     = params.get("currency_symbol", "₹")

    # Bg
    _rect(slide, 0, 0, SW, SH, LIGHT_GR)
    # Top navy bar
    _rect(slide, 0, 0, SW, Inches(1.2), NAVY)
    # Gold stripe
    _rect(slide, 0, 0, Inches(0.07), Inches(1.2), GOLD)

    # Slide label
    _text_box(slide, "EXECUTIVE SUMMARY",
              Inches(0.3), Inches(0.1), SW - Inches(0.6), Inches(0.35),
              size=8, bold=True, color=GOLD_LT)

    headline = slide_data.get("headline", params.get("title",""))
    _text_box(slide, headline, Inches(0.3), Inches(0.42),
              SW - Inches(0.6), Inches(0.65),
              size=18, bold=True, color=WHITE)

    # 3 insight cards
    insights = slide_data.get("insights", [])
    n_cards  = min(len(insights), 3)
    if n_cards:
        card_w = Inches(3.9); gap = Inches(0.22)
        total_w = n_cards * card_w + (n_cards - 1) * gap
        x_start = (SW - total_w) / 2

        for i, ins in enumerate(insights[:3]):
            cx = x_start + i * (card_w + gap)
            cy = Inches(1.4)
            ch = SH - Inches(1.7)

            # Card bg
            _rect(slide, cx, cy, card_w, ch, WHITE,
                  GOLD_LT if i == 0 else MID_GR, 1)

            # Number badge
            _rect(slide, cx, cy, Inches(0.55), Inches(0.55), GOLD)
            _text_box(slide, str(i+1), cx + Inches(0.01), cy + Inches(0.02),
                      Inches(0.53), Inches(0.5),
                      size=16, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

            # Card title
            _text_box(slide, ins.get("title",""), cx + Inches(0.65), cy + Inches(0.06),
                      card_w - Inches(0.8), Inches(0.45),
                      size=10.5, bold=True, color=NAVY)

            # Gold divider
            _rect(slide, cx + Inches(0.15), cy + Inches(0.62),
                  card_w - Inches(0.3), Inches(0.025), GOLD)

            # Body bullets
            bullets = ins.get("bullets", [ins.get("body","")[:200]])
            y_off = cy + Inches(0.75)
            for b in bullets[:6]:
                _text_box(slide, f"• {b}", cx + Inches(0.2), y_off,
                          card_w - Inches(0.35), Inches(0.4),
                          size=9, color=DARK_GR, wrap=True)
                y_off += Inches(0.5)

            # KPI at bottom of card
            kpi = ins.get("kpi", "")
            kpi_label = ins.get("kpi_label", "")
            if kpi:
                _rect(slide, cx, cy + ch - Inches(0.85),
                      card_w, Inches(0.85), NAVY)
                _text_box(slide, kpi, cx, cy + ch - Inches(0.85),
                          card_w, Inches(0.5),
                          size=18, bold=True, color=GOLD_LT,
                          align=PP_ALIGN.CENTER)
                _text_box(slide, kpi_label, cx,
                          cy + ch - Inches(0.38), card_w, Inches(0.3),
                          size=7, color=MID_GR, align=PP_ALIGN.CENTER)

    _add_slide_chrome(slide, company,
                      params.get("classification","CONFIDENTIAL"), num, total)


def _slide_kpi_dashboard(prs, slide_data, params, num, total):
    """3×2 KPI grid — navy cards with accent stripe."""
    slide = _blank_slide(prs)
    company = params.get("company_name","")

    _rect(slide, 0, 0, SW, SH, LIGHT_GR)
    _rect(slide, 0, 0, SW, Inches(0.9), NAVY)
    _rect(slide, 0, 0, Inches(0.07), Inches(0.9), GOLD)

    _text_box(slide, "KPI DASHBOARD",
              Inches(0.3), Inches(0.05), SW - Inches(0.6), Inches(0.3),
              size=8, bold=True, color=GOLD_LT)
    _text_box(slide, slide_data.get("title","Performance Metrics"),
              Inches(0.3), Inches(0.3), SW - Inches(0.6), Inches(0.52),
              size=18, bold=True, color=WHITE)

    kpis = slide_data.get("kpis", [])
    n    = min(len(kpis), 6)
    COLS = 3; ROWS = 2
    cw   = Inches(3.9); ch = Inches(2.4)
    hgap = Inches(0.22); vgap = Inches(0.2)
    x0   = Inches(0.3); y0 = Inches(1.05)

    status_colors = {
        "good":    GREEN_OK, "bad":  RED_NG,
        "warning": AMBER,    "neutral": NAVY_MID
    }

    for i, kpi in enumerate(kpis[:n]):
        col_i = i % COLS; row_i = i // COLS
        cx = x0 + col_i * (cw + hgap)
        cy = y0 + row_i * (ch + vgap)

        acc = status_colors.get(kpi.get("status","neutral"), NAVY_MID)

        # Card bg
        _rect(slide, cx, cy, cw, ch, NAVY)
        # Accent top strip
        _rect(slide, cx, cy, cw, Inches(0.13), acc)

        # Label
        _text_box(slide, kpi.get("label","").upper(),
                  cx + Inches(0.15), cy + Inches(0.18),
                  cw - Inches(0.3), Inches(0.32),
                  size=7.5, bold=True, color=GOLD_LT)

        # Value (large)
        _text_box(slide, str(kpi.get("value","—")),
                  cx, cy + Inches(0.55), cw, Inches(0.9),
                  size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        # Change
        chg = str(kpi.get("change",""))
        chg_col = GREEN_OK if "+" in chg else (RED_NG if chg.startswith("-") else GOLD_LT)
        _text_box(slide, chg, cx, cy + ch - Inches(0.7), cw, Inches(0.5),
                  size=11, bold=True, color=chg_col, align=PP_ALIGN.CENTER)

        # Sub-label
        _text_box(slide, kpi.get("sub",""),
                  cx + Inches(0.15), cy + ch - Inches(0.32),
                  cw - Inches(0.3), Inches(0.28),
                  size=7, color=MID_GR)

    _add_slide_chrome(slide, company,
                      params.get("classification","CONFIDENTIAL"), num, total)


def _slide_revenue_chart(prs, slide_data, params, num, total):
    """Bar chart slide — Revenue / Gross Profit trend."""
    slide   = _blank_slide(prs)
    company = params.get("company_name","")
    sym     = params.get("currency_symbol","₹")
    fd      = params.get("financial_data", slide_data)

    _rect(slide, 0, 0, SW, SH, LIGHT_GR)
    _rect(slide, 0, 0, SW, Inches(0.9), NAVY)
    _rect(slide, 0, 0, Inches(0.07), Inches(0.9), GOLD)

    _text_box(slide, "REVENUE PERFORMANCE",
              Inches(0.3), Inches(0.05), Inches(8), Inches(0.3),
              size=8, bold=True, color=GOLD_LT)
    _text_box(slide, slide_data.get("title","Revenue & Gross Profit Trend"),
              Inches(0.3), Inches(0.3), Inches(8), Inches(0.52),
              size=17, bold=True, color=WHITE)

    # Chart area (left 62%)
    chart_l = Inches(0.3); chart_t = Inches(1.0)
    chart_w = Inches(8.0); chart_h = Inches(5.8)

    rev    = fd.get("revenue",     [])
    gp     = fd.get("gross_profit",[])
    labels = fd.get("period_labels",[f"P{i+1}" for i in range(len(rev))])

    if rev:
        cd = ChartData()
        cd.categories = labels[:len(rev)]
        cd.add_series("Revenue",      rev)
        if gp: cd.add_series("Gross Profit", gp[:len(rev)])

        chart_shape = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            chart_l, chart_t, chart_w, chart_h, cd
        )
        ch_obj = chart_shape.chart
        ch_obj.has_legend = True
        ch_obj.has_title  = False
        ch_obj.legend.position = 3   # bottom
        ch_obj.plots[0].vary_by_categories = False
        # Color series
        from pptx.dml.color import RGBColor as RC
        if ch_obj.series:
            ch_obj.series[0].format.fill.solid()
            ch_obj.series[0].format.fill.fore_color.rgb = RC(13,27,42)
        if len(ch_obj.series) > 1:
            ch_obj.series[1].format.fill.solid()
            ch_obj.series[1].format.fill.fore_color.rgb = RC(201,168,76)
    else:
        _text_box(slide, "No revenue data provided.", chart_l, chart_t,
                  chart_w, chart_h, size=12, color=SLATE)

    # Insight panel (right 35%)
    panel_l = Inches(8.7); panel_t = Inches(1.0)
    panel_w = Inches(4.3); panel_h = Inches(5.8)
    _rect(slide, panel_l, panel_t, panel_w, panel_h, NAVY)

    _text_box(slide, "KEY HIGHLIGHTS",
              panel_l + Inches(0.2), panel_t + Inches(0.15),
              panel_w - Inches(0.3), Inches(0.35),
              size=8, bold=True, color=GOLD_LT)

    highlights = slide_data.get("highlights", [])
    if not highlights and rev:
        # Auto-derive highlights
        last = _num(rev[-1]); first = _num(rev[0])
        if first: growth = (last/first - 1)*100
        else: growth = 0
        highlights = [
            f"Period revenue: {_fmt(last, sym)}",
            f"Total period growth: {growth:.1f}%",
        ]
        if gp and rev:
            gm = _num(gp[-1])/_num(rev[-1])*100 if _num(rev[-1]) else 0
            highlights.append(f"Latest gross margin: {gm:.1f}%")

    y_hl = panel_t + Inches(0.6)
    for hl in highlights[:6]:
        _rect(slide, panel_l + Inches(0.15), y_hl,
              Inches(0.07), Inches(0.28), GOLD)
        _text_box(slide, str(hl),
                  panel_l + Inches(0.32), y_hl,
                  panel_w - Inches(0.45), Inches(0.45),
                  size=9, color=WHITE, wrap=True)
        y_hl += Inches(0.62)

    _add_slide_chrome(slide, company,
                      params.get("classification","CONFIDENTIAL"), num, total)


def _slide_pl_table(prs, slide_data, params, num, total):
    """P&L summary table slide."""
    slide   = _blank_slide(prs)
    company = params.get("company_name","")
    sym     = params.get("currency_symbol","₹")
    fd      = params.get("financial_data", slide_data)

    _rect(slide, 0, 0, SW, SH, LIGHT_GR)
    _rect(slide, 0, 0, SW, Inches(0.9), NAVY)
    _rect(slide, 0, 0, Inches(0.07), Inches(0.9), GOLD)

    _text_box(slide, "P&L SUMMARY",
              Inches(0.3), Inches(0.05), SW - Inches(0.6), Inches(0.3),
              size=8, bold=True, color=GOLD_LT)
    _text_box(slide, slide_data.get("title","Profit & Loss — Key Metrics"),
              Inches(0.3), Inches(0.3), SW - Inches(0.6), Inches(0.52),
              size=17, bold=True, color=WHITE)

    rev    = fd.get("revenue",     [])
    gp     = fd.get("gross_profit",[])
    ebitda = fd.get("ebitda",      [])
    net    = fd.get("net_profit",  [])
    gm_pct = fd.get("gp_margin",   [])
    eb_pct = fd.get("ebitda_margin",[])
    labels = fd.get("period_labels",[f"P{i+1}" for i in range(max(len(rev),1))])

    # Build python-pptx table
    n_periods = min(len(rev), 6)
    n_rows = 8; n_cols = n_periods + 1

    if n_periods == 0:
        _text_box(slide, "No financial data provided.",
                  Inches(0.5), Inches(1.5), SW - Inches(1), Inches(1),
                  size=12, color=SLATE)
        _add_slide_chrome(slide, company,
                          params.get("classification","CONFIDENTIAL"), num, total)
        return

    tbl = slide.shapes.add_table(
        n_rows, n_cols,
        Inches(0.3), Inches(1.05),
        SW - Inches(0.6), SH - Inches(1.5)
    ).table

    # Column widths
    label_w = Inches(2.5)
    data_w  = (SW - Inches(0.6) - label_w) / max(1, n_periods)
    tbl.columns[0].width = int(label_w)
    for c in range(1, n_cols):
        tbl.columns[c].width = int(data_w)

    # Row heights
    row_h = int((SH - Inches(1.5)) / n_rows)
    for r in range(n_rows):
        tbl.rows[r].height = row_h

    def cell_txt(r, c, text, bold=False, align=PP_ALIGN.RIGHT,
                 fg=DARK_GR, bg=None, size=9.5):
        cell = tbl.cell(r, c)
        cell.text = ""
        tf   = cell.text_frame
        tf.word_wrap = False
        p    = tf.paragraphs[0]
        p.alignment = align
        run  = p.add_run()
        run.text = str(text)
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = fg
        if bg:
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
        else:
            cell.fill.background()

    # Header row
    for c in range(n_cols):
        hdr = "Metric" if c == 0 else labels[c-1]
        cell_txt(0, c, hdr, bold=True, align=PP_ALIGN.CENTER,
                 fg=WHITE, bg=NAVY, size=9)

    # Data rows
    row_defs = [
        ("Revenue",          rev,    False, LIGHT_GR if False else None, False),
        ("Gross Profit",     gp,     False, None, False),
        ("  GP Margin %",    gm_pct, True,  None, True),
        ("EBITDA",           ebitda, False, None, False),
        ("  EBITDA Margin %",eb_pct, True,  None, True),
        ("Net Profit/(Loss)",net,    False, None, False),
        ("", [], False, None, False),
    ]

    for ri, (label, data, is_margin, bg_override, is_pct) in enumerate(row_defs, 1):
        if ri >= n_rows: break
        is_total = ri in (2, 4, 6)
        row_bg   = _rgb(224, 228, 235) if is_total else (
            LIGHT_GR if ri % 2 == 0 else WHITE)

        cell_txt(ri, 0, label, bold=is_total,
                 align=PP_ALIGN.LEFT, fg=NAVY if is_total else DARK_GR,
                 bg=_rgb(208,215,226) if is_total else row_bg, size=9)

        for c in range(1, n_cols):
            idx = c - 1
            raw = _num(data[idx]) if idx < len(data) else 0
            if is_pct:
                txt = f"{raw:.1f}%"
            else:
                txt = _fmt(raw, sym)
            val_col = GREEN_OK if (not is_pct and raw > 0) else (
                      RED_NG  if (not is_pct and raw < 0) else DARK_GR)
            cell_txt(ri, c, txt, bold=is_total,
                     fg=val_col if is_total else DARK_GR,
                     bg=_rgb(208,215,226) if is_total else row_bg, size=9)

    _add_slide_chrome(slide, company,
                      params.get("classification","CONFIDENTIAL"), num, total)


def _slide_cash_runway(prs, slide_data, params, num, total):
    """Cash & Runway slide with line chart."""
    slide   = _blank_slide(prs)
    company = params.get("company_name","")
    sym     = params.get("currency_symbol","₹")
    fd      = params.get("financial_data", slide_data)

    _rect(slide, 0, 0, SW, SH, LIGHT_GR)
    _rect(slide, 0, 0, SW, Inches(0.9), NAVY)
    _rect(slide, 0, 0, Inches(0.07), Inches(0.9), GOLD)

    _text_box(slide, "CASH & RUNWAY ANALYSIS",
              Inches(0.3), Inches(0.05), Inches(8), Inches(0.3),
              size=8, bold=True, color=GOLD_LT)
    _text_box(slide, slide_data.get("title","Cash Position & Runway"),
              Inches(0.3), Inches(0.3), Inches(8), Inches(0.52),
              size=17, bold=True, color=WHITE)

    cash    = fd.get("cash_balances",  [])
    labels  = fd.get("period_labels",  [f"M{i+1}" for i in range(len(cash))])
    burn    = fd.get("monthly_burn",   [])
    runway  = fd.get("runway_months",  "—")

    if cash:
        cd = ChartData()
        cd.categories = labels[:len(cash)]
        cd.add_series("Cash Balance", [_num(c) for c in cash])

        lc = slide.shapes.add_chart(
            XL_CHART_TYPE.LINE,
            Inches(0.3), Inches(1.0), Inches(8.2), Inches(5.8), cd
        )
        lch = lc.chart
        lch.has_legend = False
        lch.has_title  = False
        if lch.series:
            from pptx.dml.color import RGBColor as RC
            lch.series[0].format.line.color.rgb = RC(201,168,76)
            lch.series[0].format.line.width = Pt(2.5)
    else:
        _text_box(slide, "No cash data available.",
                  Inches(0.3), Inches(1.5), Inches(8), Inches(1),
                  size=12, color=SLATE)

    # Insight panel (right)
    px = Inches(8.7); py = Inches(1.0); pw = Inches(4.3); ph = Inches(5.8)
    _rect(slide, px, py, pw, ph, NAVY)

    _text_box(slide, "RUNWAY SUMMARY",
              px + Inches(0.2), py + Inches(0.15), pw, Inches(0.35),
              size=8, bold=True, color=GOLD_LT)

    # Large runway figure
    _text_box(slide, str(runway),
              px, py + Inches(0.65), pw, Inches(0.8),
              size=36, bold=True, color=GOLD_LT, align=PP_ALIGN.CENTER)
    _text_box(slide, "months runway",
              px, py + Inches(1.35), pw, Inches(0.35),
              size=9, color=MID_GR, align=PP_ALIGN.CENTER)

    _rect(slide, px + Inches(0.3), py + Inches(1.85),
          pw - Inches(0.6), Inches(0.03), GOLD)

    # Metrics
    metrics = [
        ("Current Cash",      fd.get("current_cash",  "—")),
        ("Monthly Burn",      fd.get("burn_rate",     "—")),
        ("Breakeven",         fd.get("breakeven",     "—")),
        ("Series A Target",   fd.get("fundraise_target","—")),
    ]
    y_m = py + Inches(2.1)
    for lbl, val in metrics:
        _text_box(slide, lbl, px + Inches(0.2), y_m, pw - Inches(0.3),
                  Inches(0.3), size=8, color=GOLD_LT)
        _text_box(slide, str(val), px + Inches(0.2), y_m + Inches(0.3),
                  pw - Inches(0.3), Inches(0.4), size=13,
                  bold=True, color=WHITE)
        y_m += Inches(0.85)

    _add_slide_chrome(slide, company,
                      params.get("classification","CONFIDENTIAL"), num, total)


def _slide_full_text(prs, slide_data, params, num, total):
    """2-column text / bullets slide."""
    slide   = _blank_slide(prs)
    company = params.get("company_name","")

    _rect(slide, 0, 0, SW, SH, LIGHT_GR)
    _rect(slide, 0, 0, SW, Inches(1.0), NAVY)
    _rect(slide, 0, 0, Inches(0.07), Inches(1.0), GOLD)

    _text_box(slide, slide_data.get("section",""), Inches(0.3), Inches(0.05),
              SW - Inches(0.6), Inches(0.3), size=8, bold=True, color=GOLD_LT)
    _text_box(slide, slide_data.get("title",""), Inches(0.3), Inches(0.3),
              SW - Inches(0.6), Inches(0.62), size=17, bold=True, color=WHITE)

    # Content
    content = slide_data.get("content","")
    bullets  = slide_data.get("bullets", [])
    if not bullets and content:
        bullets = [ln.strip("• -") for ln in content.split("\n") if ln.strip()]

    half = len(bullets) // 2 + len(bullets) % 2
    cols = [bullets[:half], bullets[half:]]

    for ci, col_bullets in enumerate(cols):
        cx = Inches(0.35) + ci * Inches(6.45)
        cy = Inches(1.15)
        for b in col_bullets:
            if not b: continue
            # Gold bullet marker
            _rect(slide, cx, cy + Inches(0.1), Inches(0.06), Inches(0.22), GOLD)
            _text_box(slide, b, cx + Inches(0.15), cy,
                      Inches(6.1), Inches(0.55),
                      size=10, color=DARK_GR, wrap=True)
            cy += Inches(0.62)

    # Source / footnote
    src = slide_data.get("source","")
    if src:
        _text_box(slide, f"Source: {src}", Inches(0.3), SH - Inches(0.6),
                  SW - Inches(0.6), Inches(0.35), size=7,
                  italic=True, color=SLATE)

    _add_slide_chrome(slide, company,
                      params.get("classification","CONFIDENTIAL"), num, total)


def _slide_next_steps(prs, slide_data, params, num, total):
    """Action items / next steps on dark bg."""
    slide   = _blank_slide(prs)
    company = params.get("company_name","")

    _rect(slide, 0, 0, SW, SH, NAVY)
    _rect(slide, 0, 0, Inches(0.07), SH, GOLD)
    _rect(slide, 0, SH - Inches(0.3), SW, Inches(0.3), GOLD)

    _text_box(slide, "NEXT STEPS & DECISIONS",
              Inches(0.3), Inches(0.2), SW - Inches(0.6), Inches(0.3),
              size=8, bold=True, color=GOLD_LT)
    _text_box(slide, slide_data.get("title","Strategic Actions Required"),
              Inches(0.3), Inches(0.5), SW - Inches(0.6), Inches(0.6),
              size=18, bold=True, color=WHITE)

    items = slide_data.get("items", [])
    n_items = min(len(items), 6)
    if n_items:
        item_h = min(Inches(0.95), (SH - Inches(1.5)) / n_items)
        for i, item in enumerate(items[:n_items]):
            iy = Inches(1.25) + i * (item_h + Inches(0.08))
            # Number badge
            _rect(slide, Inches(0.3), iy, Inches(0.55), item_h * 0.8, GOLD)
            _text_box(slide, str(i+1),
                      Inches(0.3), iy + Inches(0.06),
                      Inches(0.55), item_h * 0.7,
                      size=16, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
            # Action
            txt  = item if isinstance(item, str) else item.get("action","")
            owner = "" if isinstance(item, str) else item.get("owner","")
            by    = "" if isinstance(item, str) else item.get("timeline","")
            _text_box(slide, txt, Inches(1.0), iy + Inches(0.05),
                      Inches(8.5), Inches(0.48), size=11, bold=True, color=WHITE)
            if owner or by:
                meta = "  ".join(filter(None,[owner, by]))
                _text_box(slide, meta, Inches(1.0), iy + Inches(0.52),
                          Inches(8.5), Inches(0.35), size=8.5, color=GOLD_LT)

    _add_slide_chrome(slide, company,
                      params.get("classification","CONFIDENTIAL"), num, total)


def _slide_closing(prs, slide_data, params, num, total):
    """Thank you / closing slide."""
    slide   = _blank_slide(prs)
    company = params.get("company_name","")

    _rect(slide, 0, 0, SW, SH, NAVY)
    _rect(slide, 0, 0, Inches(0.12), SH, GOLD)
    _rect(slide, 0, SH - Inches(0.12), SW, Inches(0.12), GOLD)

    # Large "Thank You"
    _text_box(slide, "Thank You",
              Inches(1.5), SH * 0.22, SW - Inches(2), Inches(2.2),
              size=48, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    _rect(slide, Inches(2.5), SH * 0.55, SW - Inches(5), Inches(0.04), GOLD)

    contact = slide_data.get("contact","")
    website = slide_data.get("website","")

    _text_box(slide, company, Inches(0.3), SH * 0.62, SW - Inches(0.6),
              Inches(0.5), size=14, bold=True, color=GOLD_LT,
              align=PP_ALIGN.CENTER)
    if contact:
        _text_box(slide, contact, Inches(0.3), SH * 0.72, SW - Inches(0.6),
                  Inches(0.4), size=10, color=MID_GR, align=PP_ALIGN.CENTER)
    if website:
        _text_box(slide, website, Inches(0.3), SH * 0.79, SW - Inches(0.6),
                  Inches(0.35), size=10, color=GOLD_LT, align=PP_ALIGN.CENTER)

    _text_box(slide, f"Generated by OrchestrIQ | GorakhAI  ·  {params.get('classification','CONFIDENTIAL')}",
              Inches(0.3), SH - Inches(0.55), SW - Inches(0.6), Inches(0.35),
              size=7.5, color=SLATE, align=PP_ALIGN.CENTER)

    _add_slide_chrome(slide, company,
                      params.get("classification","CONFIDENTIAL"), num, total, is_title=True)


def _slide_generic(prs, slide_data, params, num, total):
    """Fallback: section + title + bulleted content."""
    slide   = _blank_slide(prs)
    company = params.get("company_name","")

    _rect(slide, 0, 0, SW, SH, LIGHT_GR)
    _rect(slide, 0, 0, SW, Inches(1.0), NAVY)
    _rect(slide, 0, 0, Inches(0.07), Inches(1.0), GOLD)

    _text_box(slide, slide_data.get("section",""), Inches(0.3), Inches(0.06),
              SW - Inches(0.6), Inches(0.28), size=7.5, bold=True, color=GOLD_LT)
    _text_box(slide, slide_data.get("title",""), Inches(0.3), Inches(0.32),
              SW - Inches(0.6), Inches(0.6), size=17, bold=True, color=WHITE)

    body = slide_data.get("content","")
    bullets = slide_data.get("bullets", [])
    if not bullets and body:
        bullets = [ln.strip("• -") for ln in body.split("\n") if ln.strip()]

    y = Inches(1.12)
    for b in bullets[:12]:
        if not b.strip(): continue
        _rect(slide, Inches(0.32), y + Inches(0.12), Inches(0.06), Inches(0.22), GOLD)
        _text_box(slide, b.strip(), Inches(0.5), y,
                  SW - Inches(0.85), Inches(0.48),
                  size=10, color=DARK_GR, wrap=True)
        y += Inches(0.52)

    _add_slide_chrome(slide, company,
                      params.get("classification","CONFIDENTIAL"), num, total)

# ─── SLIDE DISPATCHER ─────────────────────────────────────────────────────────

SLIDE_MAP = {
    "title":           _slide_title,
    "exec_summary":    _slide_exec_summary,
    "kpi_dashboard":   _slide_kpi_dashboard,
    "revenue_chart":   _slide_revenue_chart,
    "pl_table":        _slide_pl_table,
    "cash_runway":     _slide_cash_runway,
    "full_text":       _slide_full_text,
    "next_steps":      _slide_next_steps,
    "closing":         _slide_closing,
}

# ─── MAIN BUILDER ─────────────────────────────────────────────────────────────

def build_pptx(params: Dict[str, Any]) -> bytes:
    prs = Presentation()
    prs.slide_width  = SW
    prs.slide_height = SH

    slides_spec = params.get("slides", [])

    # Ensure title + closing bookends if not in spec
    if not slides_spec or slides_spec[0].get("layout") != "title":
        slides_spec = [{"layout": "title"}] + slides_spec
    if not slides_spec or slides_spec[-1].get("layout") != "closing":
        slides_spec.append({
            "layout":  "closing",
            "contact": params.get("contact",""),
            "website": params.get("website",""),
        })

    total = len(slides_spec)

    for i, sd in enumerate(slides_spec):
        layout = sd.get("layout","generic")
        fn     = SLIDE_MAP.get(layout, _slide_generic)
        num    = i + 1

        if layout == "title":
            fn(prs, params)
        elif layout == "closing":
            fn(prs, sd, params, num, total)
        else:
            fn(prs, sd, params, num, total)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()
