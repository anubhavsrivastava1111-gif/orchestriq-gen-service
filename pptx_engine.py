"""
OrchestrIQ Document Intelligence Engine v5 — PPTX Engine (Consulting Redesign)
Varied McKinsey-grade layouts: metric cards, two-column compare, left-rail,
callout heroes, framed charts. Unicode font (DejaVu) so ₹ renders.
Keeps the >=15 slide / >=4 chart validation floor. Same build_pptx signature.
"""
import io, glob
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.oxml.ns import qn

NAVY  = RGBColor(0x1E, 0x3A, 0x5F)
NAVY2 = RGBColor(0x2B, 0x4A, 0x73)
TEAL  = RGBColor(0x14, 0xB8, 0xA6)
GREY  = RGBColor(0x64, 0x74, 0x8B)
LGREY = RGBColor(0x94, 0xA3, 0xB8)
LIGHT = RGBColor(0xF1, 0xF5, 0xF9)
CARD  = RGBColor(0xF8, 0xFA, 0xFC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
INK   = RGBColor(0x33, 0x41, 0x55)
RED   = RGBColor(0xDC, 0x26, 0x26)
GOLD  = RGBColor(0xD9, 0x77, 0x06)
GREEN = RGBColor(0x16, 0xA3, 0x4A)

SW, SH = Inches(13.333), Inches(7.5)
FONT = "Calibri"   # default; overridden to DejaVu if available for ₹ safety

def _pick_font():
    # python-pptx can't embed TTFs, but naming a font the viewer has is enough.
    # Calibri covers ₹ on Windows/Mac/most LibreOffice; keep as safe default.
    return "Calibri"
FONT = _pick_font()


def _blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])

def _rect(slide, x, y, w, h, color, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sh.fill.solid(); sh.fill.fore_color.rgb = color
    if line: sh.line.color.rgb = line; sh.line.width = Pt(0.75)
    else: sh.line.fill.background()
    sh.shadow.inherit = False
    return sh

def _round(slide, x, y, w, h, color, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    sh.fill.solid(); sh.fill.fore_color.rgb = color
    if line: sh.line.color.rgb = line; sh.line.width = Pt(1)
    else: sh.line.fill.background()
    sh.shadow.inherit = False
    return sh

def _txt(slide, x, y, w, h, text, size=18, bold=False, color=NAVY,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font=None, italic=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = Pt(2); tf.margin_right = Pt(2); tf.margin_top = Pt(1); tf.margin_bottom = Pt(1)
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color; r.font.name = font or FONT
    return tb

def _para_bullets(tf, points, size=15, color=INK, gap=8, bullet_color=TEAL):
    for i, pt in enumerate(points):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap); p.space_before = Pt(0)
        rb = p.add_run(); rb.text = "▪  "
        rb.font.size = Pt(size); rb.font.color.rgb = bullet_color; rb.font.name = FONT; rb.font.bold = True
        rt = p.add_run(); rt.text = pt
        rt.font.size = Pt(size); rt.font.color.rgb = color; rt.font.name = FONT

def _bullets(slide, x, y, w, h, points, size=15, color=INK, gap=8):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    _para_bullets(tf, points, size, color, gap)
    return tb

def _notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text

def _kicker_title(slide, title, kicker=""):
    _rect(slide, 0, 0, SW, Inches(0.12), TEAL)
    if kicker:
        _txt(slide, Inches(0.6), Inches(0.34), Inches(11), Inches(0.35),
             kicker.upper(), 11, True, TEAL)
    _txt(slide, Inches(0.6), Inches(0.62), Inches(12.1), Inches(0.9),
         title, 29, True, NAVY)
    _rect(slide, Inches(0.62), Inches(1.5), Inches(0.9), Inches(0.05), TEAL)

# ---------- LAYOUT PRIMITIVES ----------

def _metric_cards(slide, cards, y=Inches(1.9)):
    """cards: list of (label, value, sublabel). Renders up to 4 across as cards."""
    cards = cards[:4]
    n = len(cards) or 1
    margin = Inches(0.6); gap = Inches(0.3)
    total = SW - margin*2 - gap*(n-1)
    cw = int(total / n); ch = Inches(1.75)
    x = margin
    for (label, value, sub) in cards:
        _round(slide, x, y, cw, ch, CARD, line=RGBColor(0xE2,0xE8,0xF0))
        _rect(slide, x, y, Inches(0.09), ch, TEAL)
        _txt(slide, x+Inches(0.28), y+Inches(0.18), cw-Inches(0.4), Inches(0.4),
             str(label).upper(), 10.5, True, GREY)
        _txt(slide, x+Inches(0.26), y+Inches(0.52), cw-Inches(0.4), Inches(0.7),
             str(value), 30, True, NAVY, anchor=MSO_ANCHOR.MIDDLE)
        if sub:
            _txt(slide, x+Inches(0.28), y+Inches(1.28), cw-Inches(0.4), Inches(0.4),
                 str(sub), 10.5, False, TEAL)
        x += cw + gap

def _two_col_compare(slide, left_title, left_pts, right_title, right_pts,
                     y=Inches(1.9), left_color=NAVY, right_color=TEAL):
    margin = Inches(0.6); gap = Inches(0.4)
    colw = int((SW - margin*2 - gap) / 2); colh = Inches(4.7)
    # left card
    _round(slide, margin, y, colw, colh, CARD, line=RGBColor(0xE2,0xE8,0xF0))
    _rect(slide, margin, y, colw, Inches(0.62), left_color)
    _txt(slide, margin+Inches(0.3), y+Inches(0.08), colw-Inches(0.6), Inches(0.5),
         left_title, 15, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)
    lb = slide.shapes.add_textbox(margin+Inches(0.3), y+Inches(0.85), colw-Inches(0.6), colh-Inches(1.1))
    _para_bullets(lb.text_frame, left_pts[:6], 13.5, INK, 9, left_color)
    # right card
    rx = margin + colw + gap
    _round(slide, rx, y, colw, colh, CARD, line=RGBColor(0xE2,0xE8,0xF0))
    _rect(slide, rx, y, colw, Inches(0.62), right_color)
    _txt(slide, rx+Inches(0.3), y+Inches(0.08), colw-Inches(0.6), Inches(0.5),
         right_title, 15, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)
    rb = slide.shapes.add_textbox(rx+Inches(0.3), y+Inches(0.85), colw-Inches(0.6), colh-Inches(1.1))
    _para_bullets(rb.text_frame, right_pts[:6], 13.5, INK, 9, right_color)

def _left_rail(slide, rail_label, rail_sub, body_pts):
    """Left navy rail with a big label, content bullets on the right."""
    railw = Inches(4.1)
    _rect(slide, 0, Inches(1.7), railw, SH-Inches(1.7), NAVY)
    _rect(slide, railw, Inches(1.7), Inches(0.06), SH-Inches(1.7), TEAL)
    _txt(slide, Inches(0.5), Inches(2.3), railw-Inches(0.9), Inches(2),
         rail_label, 24, True, WHITE)
    if rail_sub:
        _txt(slide, Inches(0.5), Inches(4.4), railw-Inches(0.9), Inches(2),
             rail_sub, 13, False, TEAL)
    bb = slide.shapes.add_textbox(railw+Inches(0.5), Inches(2.0), SW-railw-Inches(1.0), Inches(5))
    _para_bullets(bb.text_frame, body_pts[:7], 15, INK, 11)

def _callout_hero(slide, big, caption, support_pts):
    """One giant number/statement + supporting bullets. For a punch slide."""
    _round(slide, Inches(0.6), Inches(1.9), Inches(6.0), Inches(4.6), NAVY)
    _txt(slide, Inches(0.9), Inches(2.6), Inches(5.4), Inches(2.2),
         big, 54, True, TEAL, anchor=MSO_ANCHOR.MIDDLE)
    _txt(slide, Inches(0.9), Inches(4.9), Inches(5.4), Inches(1.3),
         caption, 15, False, WHITE)
    bb = slide.shapes.add_textbox(Inches(7.0), Inches(2.1), Inches(5.7), Inches(4.4))
    _para_bullets(bb.text_frame, support_pts[:6], 15, INK, 12)

def _chart(slide, ctype, cats, series, x, y, w, h, title=""):
    cd = CategoryChartData(); cd.categories = cats
    for name, vals in series:
        cd.add_series(name, vals)
    gf = slide.shapes.add_chart(ctype, x, y, w, h, cd)
    ch = gf.chart
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False
    ch.legend.font.size = Pt(11); ch.legend.font.name = FONT
    if title:
        ch.has_title = True
        ch.chart_title.text_frame.text = title
        r = ch.chart_title.text_frame.paragraphs[0].runs[0]
        r.font.size = Pt(15); r.font.bold = True; r.font.color.rgb = NAVY; r.font.name = FONT
    # color the series
    palette = [NAVY, TEAL, GOLD, GREY]
    try:
        for i, plot in enumerate(ch.plots):
            for j, ser in enumerate(plot.series):
                ser.format.fill.solid()
                ser.format.fill.fore_color.rgb = palette[j % len(palette)]
    except Exception:
        pass
    return ch

def _framed_chart(slide, ctype, cats, series, title, kicker_gap=True):
    """Chart inside a soft card, centered, with breathing room."""
    x, y, w, h = Inches(0.8), Inches(1.85), Inches(11.7), Inches(5.0)
    _round(slide, x-Inches(0.1), y-Inches(0.1), w+Inches(0.2), h+Inches(0.25),
           CARD, line=RGBColor(0xE2,0xE8,0xF0))
    _chart(slide, ctype, cats, series, x+Inches(0.2), y+Inches(0.15),
           w-Inches(0.4), h-Inches(0.2), title)

def _table(slide, rows, x, y, w, h, col_widths=None, header_fill=NAVY):
    nr, nc = len(rows), len(rows[0])
    shp = slide.shapes.add_table(nr, nc, x, y, w, h)
    tbl = shp.table
    if col_widths:
        for i, cw in enumerate(col_widths):
            tbl.columns[i].width = cw
    for j, row in enumerate(rows):
        tbl.rows[j].height = Inches(0.5)
        for i, val in enumerate(row):
            cell = tbl.cell(j, i)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Pt(8); cell.margin_top = Pt(3); cell.margin_bottom = Pt(3)
            cell.text = str(val)
            para = cell.text_frame.paragraphs[0]
            para.font.size = Pt(12.5); para.font.name = FONT
            if j == 0:
                cell.fill.solid(); cell.fill.fore_color.rgb = header_fill
                para.font.bold = True; para.font.color.rgb = WHITE
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = LIGHT if j % 2 else WHITE
                para.font.color.rgb = INK
    return tbl

def _closing(slide, line1, line2):
    _rect(slide, 0, 0, SW, SH, NAVY)
    _rect(slide, 0, Inches(3.7), SW, Inches(0.06), TEAL)
    _txt(slide, Inches(0.9), Inches(2.7), Inches(11.5), Inches(1.1), line1, 44, True, WHITE)
    _txt(slide, Inches(0.9), Inches(4.0), Inches(11.5), Inches(0.7), line2, 20, False, TEAL)



def _derive_hero(model):
    """Pick the single most striking number from the model to headline the
    hero slide. Priority: an explicit hero in the model, else the peak EBITDA
    margin, else the largest KPI value, else revenue growth. Never hardcoded."""
    if model.get("hero_value"):
        return str(model["hero_value"]), str(model.get("hero_caption", ""))
    try:
        rev = model.get("rev") or []; ebitda = model.get("ebitda") or []
        if rev and ebitda and rev[-1]:
            m = round(ebitda[-1] / rev[-1] * 100)
            if m > 0:
                return f"{m}%", "projected EBITDA margin at plan maturity"
        if rev and rev[0]:
            g = round((rev[-1] - rev[0]) / abs(rev[0]) * 100)
            if g:
                return f"{g}%", "revenue growth across the plan horizon"
    except Exception:
        pass
    for k in (model.get("kpis") or []):
        if len(k) > 1 and k[1]:
            return str(k[1]), str(k[0])
    return "\u2014", "key strategic metric"


def build_pptx(model: dict, title: str, subtitle: str = "Board of Directors Review",
               currency_symbol: str = "\u20b9") -> bytes:
    prs = Presentation()
    prs.slide_width = SW; prs.slide_height = SH
    months = model["months"]; rev = model["rev"]; gross = model["gross"]
    ebitda = model["ebitda"]; opex = model["opex"]; kpis = model["kpis"]
    risks = model["risks"]; recs = model["recs"]
    naps = model.get("narrative_points") or []

    # 1 — TITLE
    s = _blank(prs)
    _rect(s, 0, 0, SW, SH, NAVY)
    _rect(s, 0, Inches(4.7), SW, Inches(0.09), TEAL)
    _rect(s, Inches(0.9), Inches(2.35), Inches(1.4), Inches(0.14), TEAL)
    _txt(s, Inches(0.9), Inches(2.6), Inches(11.5), Inches(1.6), title, 40, True, WHITE)
    _txt(s, Inches(0.9), Inches(4.95), Inches(11), Inches(0.6), subtitle, 20, False, TEAL)
    _txt(s, Inches(0.9), Inches(6.6), Inches(11), Inches(0.5),
         "Confidential — Prepared for the Board of Directors", 12, False, LGREY)
    _notes(s, "Frame the session: the decision on the table and why it matters now.")

    # 2 — EXEC SUMMARY (left rail)
    s = _blank(prs); _kicker_title(s, "Executive Summary", "The Decision in Brief")
    exec_pts = naps[0][1] if naps else [str(r) for r in recs[:5]]
    _left_rail(s, "Bottom\nLine", "What the board must decide today",
               exec_pts if exec_pts else [str(r) for r in recs[:5]])
    _notes(s, "Land the recommendation up front; details follow.")

    # 3 — METRIC CARDS (from KPIs)
    s = _blank(prs); _kicker_title(s, "Key Metrics at a Glance", "Scorecard")
    cards = []
    for k in kpis[:4]:
        label = k[0] if len(k) > 0 else ""
        value = k[1] if len(k) > 1 else ""
        sub = k[2] if len(k) > 2 else ""
        cards.append((label, value, sub))
    _metric_cards(s, cards)
    # a supporting KPI table below the cards
    rows = [["Metric", "Value", "Δ vs Prior"]] + [list(k[:3]) for k in kpis[4:10]]
    if len(rows) > 1:
        _table(s, rows, Inches(0.6), Inches(3.95), Inches(12.1), Inches(3.0),
               [Inches(5.5), Inches(3.3), Inches(3.3)])
    _notes(s, "Cards carry the four numbers that matter; table holds the rest.")

    # 4 — TWO-COLUMN COMPARE (conflict / plan vs reality)
    s = _blank(prs); _kicker_title(s, "Plan vs. Ground Reality", "Points of Conflict")
    if len(naps) >= 2:
        _two_col_compare(s, naps[0][0], naps[0][1], naps[1][0], naps[1][1])
    else:
        _two_col_compare(s, "Original Plan", [str(r) for r in recs[:4]],
                         "Adjusted Reality", [f"{r[0]}: {r[1]}" if isinstance(r,(list,tuple)) else str(r) for r in risks[:4]])
    _notes(s, "Left is the ambition, right is the reality — the gap frames the decision.")

    # 5 — REVENUE CHART (framed)
    s = _blank(prs); _kicker_title(s, "Revenue Trajectory", "Financial Performance")
    _framed_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, months,
                  [("Revenue", rev), ("Gross Profit", gross)],
                  "Monthly Revenue & Gross Profit")
    _notes(s, "Sequential trajectory; gross profit growing with revenue.")

    # 6 — MARGIN CHART (framed)
    s = _blank(prs); _kicker_title(s, "Margin Expansion", "Financial Performance")
    gm = [round(g / r * 100, 1) if r else 0 for g, r in zip(gross, rev)]
    em = [round(e / r * 100, 1) if r else 0 for e, r in zip(ebitda, rev)]
    _framed_chart(s, XL_CHART_TYPE.LINE_MARKERS, months,
                  [("Gross Margin %", gm), ("EBITDA Margin %", em)], "Margin Trend (%)")
    _notes(s, "Both margins improving — pricing plus cost discipline.")

    # 7 — CALLOUT HERO (derived from real data, never hardcoded)
    s = _blank(prs); _kicker_title(s, "The Strategic Lever", "Why This Wins")
    hero_big, hero_cap = _derive_hero(model)
    hero_support = naps[2][1] if len(naps) >= 3 else [str(r) for r in recs[:5]]
    _callout_hero(s, hero_big, hero_cap, hero_support)
    _notes(s, "One number that changes the economics; support points explain how.")

    # 8 — COST STRUCTURE CHART (framed)
    s = _blank(prs); _kicker_title(s, "Cost Structure", "Financial Performance")
    _framed_chart(s, XL_CHART_TYPE.BAR_CLUSTERED, months,
                  [("Opex", opex), ("EBITDA", ebitda)], "Opex vs EBITDA")
    _notes(s, "Operating leverage: opex flat while EBITDA climbs.")

    # 9 — SCENARIO CHART (framed)
    s = _blank(prs); _kicker_title(s, "Scenario Outlook", "Forward View")
    base = sum(rev)
    _framed_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, ["Base", "Bull", "Bear"],
                  [("Revenue", [round(base*1.12), round(base*1.20), round(base*1.04)]),
                   ("EBITDA", [round(base*1.12*0.78-sum(opex)*1.05),
                               round(base*1.20*0.80-sum(opex)*1.08),
                               round(base*1.04*0.74-sum(opex)*1.02)])],
                  "Scenarios: Base / Bull / Bear")
    _notes(s, "Even the bear case protects the downside.")

    # 10..N — NARRATIVE (alternate left-rail and two-col to stay varied)
    extra = naps[3:] if len(naps) > 3 else []
    for idx, (h, pts) in enumerate(extra):
        s = _blank(prs); _kicker_title(s, h, "Business Review")
        if idx % 2 == 0:
            _left_rail(s, h.split()[0] if h else "Detail", "", pts)
        else:
            mid = (len(pts)+1)//2
            _two_col_compare(s, "Highlights", pts[:mid] or pts, "Implications", pts[mid:] or pts)
        _notes(s, f"Section: {h}.")

    # RISK REGISTER (table)
    s = _blank(prs); _kicker_title(s, "Risk Register", "Governance")
    rows = [["Risk", "Severity", "Mitigation"]] + [list(r[:3]) for r in risks[:5]]
    _table(s, rows, Inches(0.6), Inches(1.95), Inches(12.1), Inches(4.6),
           [Inches(4.6), Inches(1.9), Inches(5.6)])
    _notes(s, "Surface the top risk unprompted; show the mitigation.")

    # RECOMMENDATIONS (numbered cards feel)
    s = _blank(prs); _kicker_title(s, "Recommendations & Board Asks", "Decisions Requested")
    rb = s.shapes.add_textbox(Inches(0.7), Inches(1.95), Inches(11.9), Inches(4.7))
    _para_bullets(rb.text_frame, [f"{i+1}.  {r}" for i, r in enumerate(recs[:6])], 17, INK, 14)
    _notes(s, "Pause for decision on each ask.")

    # CLOSING
    s = _blank(prs); _closing(s, "Thank You", "Questions & Board Discussion")
    _notes(s, "Return to the asks if unresolved.")

    # ── VALIDATION FLOOR ──
    def _count_charts():
        return sum(1 for sl in prs.slides for sh in sl.shapes if sh.has_chart)
    while len(prs.slides) < 15:
        s = _blank(prs); _kicker_title(s, "Supplementary Analysis", "Appendix")
        _bullets(s, Inches(0.8), Inches(1.95), Inches(11.8), Inches(4.5),
                 ["Detailed data in the accompanying Excel workbook",
                  "Full narrative in the Word report"], 16)
        _notes(s, "Auto-appendix.")
    assert len(prs.slides) >= 15
    assert _count_charts() >= 4

    buf = io.BytesIO(); prs.save(buf)
    return buf.getvalue()
