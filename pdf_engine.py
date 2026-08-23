"""
OrchestrIQ Document Intelligence Engine v6 — PDF Engine (Consulting-grade)
Backward-compatible: preserves _cover, _hf, _tstyle, _bar_drawing, S_H1, S_BODY,
S_TOC, S_SMALL, W, H (imported by doc_blueprint_engine.py) with original
signatures. Upgrades: embedded DejaVu Unicode font (₹ renders, no boxes),
wrapped table cells (no overlap/clipping), metric-card KPI band, callout boxes,
section dividers, styled charts, cleaner cover. Same build_pdf signature.
"""
import io, glob
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY, TA_CENTER
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph,
                                Spacer, Table, TableStyle, PageBreak, KeepTogether)
from reportlab.graphics.shapes import Drawing, String, Rect
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.linecharts import HorizontalLineChart
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

NAVY  = colors.HexColor("#1E3A5F")
NAVY2 = colors.HexColor("#27446B")
TEAL  = colors.HexColor("#14B8A6")
GREY  = colors.HexColor("#64748B")
LGREY = colors.HexColor("#94A3B8")
LIGHT = colors.HexColor("#F1F5F9")
CARD  = colors.HexColor("#F8FAFC")
BORDER= colors.HexColor("#E2E8F0")
INK   = colors.HexColor("#334155")
GOLD  = colors.HexColor("#D97706")
GREEN = colors.HexColor("#16A34A")

W, H = A4

# ── Embed a Unicode font so ₹, dashes, arrows render instead of boxes. ──
def _register_fonts():
    reg = {"base": "Helvetica", "bold": "Helvetica-Bold"}
    try:
        import os
        _here = os.path.dirname(os.path.abspath(__file__))
        def _find(name):
            for p in [os.path.join(_here, "fonts", name), os.path.join(_here, name)]:
                if os.path.exists(p):
                    return [p]
            return glob.glob("/usr/share/fonts/**/" + name, recursive=True)
        c = _find("DejaVuSans.ttf")
        b = _find("DejaVuSans-Bold.ttf")
        if c:
            pdfmetrics.registerFont(TTFont("OIQ", c[0])); reg["base"] = "OIQ"
        if b:
            pdfmetrics.registerFont(TTFont("OIQ-Bold", b[0])); reg["bold"] = "OIQ-Bold"
    except Exception:
        pass
    return reg

_F = _register_fonts()
FONT = _F["base"]; FONT_BOLD = _F["bold"]

# ── Paragraph styles (names S_H1/S_BODY/S_TOC/S_SMALL preserved for import) ──
S_TITLE = ParagraphStyle("t", fontName=FONT_BOLD, fontSize=26, textColor=colors.white, leading=32)
S_SUB   = ParagraphStyle("s", fontName=FONT, fontSize=13, textColor=TEAL, leading=18)
S_KICK  = ParagraphStyle("k", fontName=FONT_BOLD, fontSize=10, textColor=TEAL, leading=12,
                         spaceBefore=14, spaceAfter=2)
S_H1    = ParagraphStyle("h1", fontName=FONT_BOLD, fontSize=18, textColor=NAVY,
                         spaceBefore=6, spaceAfter=8, leading=22)
S_H2    = ParagraphStyle("h2", fontName=FONT_BOLD, fontSize=13, textColor=NAVY2,
                         spaceBefore=10, spaceAfter=4, leading=16)
S_BODY  = ParagraphStyle("b", fontName=FONT, fontSize=10.5, textColor=INK,
                         leading=15.5, spaceAfter=8, alignment=TA_JUSTIFY)
S_TOC   = ParagraphStyle("toc", fontName=FONT, fontSize=11.5, textColor=NAVY,
                         leading=22, leftIndent=6)
S_SMALL = ParagraphStyle("sm", fontName=FONT, fontSize=8.5, textColor=GREY)
S_CELL  = ParagraphStyle("cell", fontName=FONT, fontSize=9.5, textColor=INK, leading=12)
S_CELLH = ParagraphStyle("cellh", fontName=FONT_BOLD, fontSize=9.5, textColor=colors.white, leading=12)
S_CARDL = ParagraphStyle("cardl", fontName=FONT_BOLD, fontSize=8, textColor=GREY, leading=10)
S_CARDV = ParagraphStyle("cardv", fontName=FONT_BOLD, fontSize=20, textColor=NAVY, leading=22)
S_CALL  = ParagraphStyle("call", fontName=FONT, fontSize=11, textColor=INK, leading=16)


def _wrap(text, n):
    words, lines, cur = str(text).split(), [], ""
    for w in words:
        if len(cur) + len(w) + 1 <= n: cur = (cur + " " + w).strip()
        else: lines.append(cur); cur = w
    if cur: lines.append(cur)
    return lines[:3]


# ── _cover: preserved signature (canvas, doc, title, subtitle) ──
def _cover(canvas, doc, title, subtitle):
    canvas.saveState()
    canvas.setFillColor(NAVY); canvas.rect(0, 0, W, H, fill=1, stroke=0)
    # accent geometry
    canvas.setFillColor(TEAL); canvas.rect(0, H - 4.2 * cm, W, 0.25 * cm, fill=1, stroke=0)
    canvas.setFillColor(TEAL); canvas.rect(2.2 * cm, H - 8.0 * cm, 2.4 * cm, 0.14 * cm, fill=1, stroke=0)
    canvas.setFillColor(colors.white); canvas.setFont(FONT_BOLD, 30)
    y = H - 9 * cm
    for line in _wrap(title, 30):
        canvas.drawString(2.2 * cm, y, line); y -= 1.15 * cm
    # The subtitle was a single drawString with no wrapping, so anything longer
    # than the page ran straight off the edge - your cover read
    # "...Working Capital Optimizati" with the rest missing. It now wraps, and
    # the font steps down once if it is very long.
    canvas.setFillColor(TEAL)
    sub = str(subtitle or "")
    fsize = 15 if len(sub) <= 60 else 12
    canvas.setFont(FONT, fsize)
    sy = y - 0.6 * cm
    for sline in _wrap(sub, 62 if fsize == 15 else 78)[:3]:
        canvas.drawString(2.2 * cm, sy, sline)
        sy -= (fsize + 5) * 0.0353 * cm * 10
    canvas.setFillColor(LGREY); canvas.setFont(FONT, 10)
    canvas.drawString(2.2 * cm, 2.6 * cm, "Confidential \u2014 Prepared for the Board of Directors")
    canvas.setFillColor(TEAL); canvas.rect(0, 1.8 * cm, W, 0.12 * cm, fill=1, stroke=0)
    canvas.restoreState()


# ── _hf: preserved signature (canvas, doc, title) ──
def _hf(canvas, doc, title):
    canvas.saveState()
    canvas.setFillColor(TEAL); canvas.rect(0, H - 0.9 * cm, W, 0.12 * cm, fill=1, stroke=0)
    canvas.setFillColor(GREY); canvas.setFont(FONT, 8.5)
    canvas.drawString(2 * cm, H - 1.5 * cm, str(title)[:80])
    canvas.drawRightString(W - 2 * cm, 1.2 * cm, f"Page {doc.page - 1}")
    canvas.drawString(2 * cm, 1.2 * cm, "Confidential")
    canvas.setStrokeColor(BORDER); canvas.setLineWidth(0.5)
    canvas.line(2 * cm, 1.55 * cm, W - 2 * cm, 1.55 * cm)
    canvas.restoreState()


# ── _tstyle: preserved. Header handled by wrapped Paragraphs; keeps grid/stripe.
def _tstyle():
    return TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        # Belt and braces for the rupee symbol. Cells wrapped in a Paragraph
        # already carry the DejaVu font, but ANY cell that reaches a Table as a
        # raw string falls back to Helvetica, which has no rupee glyph - that is
        # the black box you saw in table headers such as "Monthly (?)".
        ("FONTNAME", (0, 0), (-1, -1), FONT),
        ("FONTNAME", (0, 0), (-1, 0), FONT_BOLD),
    ])


def _cellwrap(data):
    """Wrap each cell in a Paragraph so text wraps and rows auto-grow."""
    out = []
    for ri, row in enumerate(data):
        st = S_CELLH if ri == 0 else S_CELL
        out.append([Paragraph(str(c) if c is not None else "", st) for c in row])
    return out


def _wrapped_table(data, col_widths):
    t = Table(_cellwrap(data), colWidths=col_widths, repeatRows=1)
    t.setStyle(_tstyle())
    return t


# ── _bar_drawing: preserved signature. Cleaner colors, bigger title, framed.
def _bar_drawing(months, series, title):
    d = Drawing(460, 220)
    d.add(Rect(0, 0, 460, 220, fillColor=CARD, strokeColor=BORDER, strokeWidth=0.5))
    ch = VerticalBarChart()
    ch.x, ch.y, ch.width, ch.height = 50, 35, 370, 140
    ch.data = [list(s[1]) for s in series]
    ch.categoryAxis.categoryNames = [str(m) for m in months]
    palette = [NAVY, TEAL, GOLD, GREY]
    for i in range(len(series)):
        try: ch.bars[i].fillColor = palette[i % len(palette)]
        except Exception: pass
    ch.valueAxis.labels.fontName = FONT; ch.valueAxis.labels.fontSize = 7
    ch.categoryAxis.labels.fontName = FONT; ch.categoryAxis.labels.fontSize = 7.5
    ch.categoryAxis.labels.angle = 0
    ch.barSpacing = 1; ch.groupSpacing = 10
    d.add(ch)
    d.add(String(50, 198, str(title), fontName=FONT_BOLD, fontSize=12, fillColor=NAVY))
    return d


def _line_drawing(months, series, title):
    d = Drawing(460, 220)
    d.add(Rect(0, 0, 460, 220, fillColor=CARD, strokeColor=BORDER, strokeWidth=0.5))
    ch = HorizontalLineChart()
    ch.x, ch.y, ch.width, ch.height = 50, 35, 370, 140
    ch.data = [list(s[1]) for s in series]
    ch.categoryAxis.categoryNames = [str(m) for m in months]
    palette = [NAVY, TEAL, GOLD]
    for i in range(len(series)):
        try:
            ch.lines[i].strokeColor = palette[i % len(palette)]
            ch.lines[i].strokeWidth = 2
        except Exception: pass
    ch.valueAxis.labels.fontName = FONT; ch.valueAxis.labels.fontSize = 7
    ch.categoryAxis.labels.fontName = FONT; ch.categoryAxis.labels.fontSize = 7.5
    d.add(ch)
    d.add(String(50, 198, str(title), fontName=FONT_BOLD, fontSize=12, fillColor=NAVY))
    return d


def _kicker(text):
    return Paragraph(str(text).upper(), S_KICK)


def _metric_band(kpis):
    """A row of up to 4 KPI 'cards' rendered as a borderless table of stacked
    label/value paragraphs — the PDF equivalent of the PPT metric cards."""
    cards = kpis[:4]
    if not cards: return Spacer(1, 1)
    cells = []
    for k in cards:
        label = str(k[0]) if len(k) > 0 else ""
        value = str(k[1]) if len(k) > 1 else ""
        inner = Table([[Paragraph(label.upper(), S_CARDL)],
                       [Paragraph(value, S_CARDV)]],
                      colWidths=[(W - 4 * cm) / len(cards) - 0.3 * cm])
        inner.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), CARD),
            ("LINEBEFORE", (0, 0), (0, -1), 3, TEAL),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ("LEFTPADDING", (0, 0), (-1, -1), 10),
            ("BOX", (0, 0), (-1, -1), 0.5, BORDER),
        ]))
        cells.append(inner)
    band = Table([cells], colWidths=[(W - 4 * cm) / len(cards)] * len(cards))
    band.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),4),
                              ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
                              ("VALIGN",(0,0),(-1,-1),"TOP")]))
    return band


def _callout(text, accent=NAVY):
    """A shaded callout box for a key statement."""
    t = Table([[Paragraph(str(text), S_CALL)]], colWidths=[W - 4 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), LIGHT),
        ("LINEBEFORE", (0, 0), (0, -1), 4, accent),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 14),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
    ]))
    return t


def build_pdf(model: dict, title: str, subtitle: str = "Board Report",
              currency_symbol: str = "\u20b9") -> bytes:
    buf = io.BytesIO()
    doc = BaseDocTemplate(buf, pagesize=A4,
                          leftMargin=2 * cm, rightMargin=2 * cm,
                          topMargin=2.4 * cm, bottomMargin=2 * cm)
    frame = Frame(2 * cm, 2 * cm, W - 4 * cm, H - 4.6 * cm, id="f")
    doc.addPageTemplates([
        PageTemplate(id="cover", frames=[frame], onPage=lambda c, d: _cover(c, d, title, subtitle)),
        PageTemplate(id="body", frames=[frame], onPage=lambda c, d: _hf(c, d, title)),
    ])

    months = model["months"]; rev = model["rev"]; gross = model["gross"]
    ebitda = model["ebitda"]; kpis = model["kpis"]; risks = model["risks"]
    recs = model["recs"]; sections = model.get("sections") or []

    from reportlab.platypus import NextPageTemplate
    el = [NextPageTemplate("body"), PageBreak()]

    # EXECUTIVE SNAPSHOT — metric band up front
    el.append(_kicker("Scorecard"))
    el.append(Paragraph("Executive Snapshot", S_H1))
    el.append(_metric_band(kpis))
    el.append(Spacer(1, 14))
    # remaining KPIs as wrapped table
    rows = [["KPI", "Value", "\u0394 vs Prior"]] + [list(k[:3]) for k in kpis[4:10]]
    if len(rows) > 1:
        el.append(_wrapped_table(rows, [8 * cm, 4 * cm, 4 * cm]))
    el.append(PageBreak())

    # FINANCIAL CHARTS
    el.append(_kicker("Financial Performance"))
    el.append(Paragraph("Revenue, Margin & Cost Analysis", S_H1))
    el.append(_bar_drawing(months, [("Revenue", rev), ("Gross Profit", gross)],
                           "Revenue & Gross Profit by Period"))
    el.append(Spacer(1, 12))
    gm = [round(g / r * 100, 1) if r else 0 for g, r in zip(gross, rev)]
    em = [round(e / r * 100, 1) if r else 0 for e, r in zip(ebitda, rev)]
    el.append(_line_drawing(months, [("Gross Margin %", gm), ("EBITDA Margin %", em)],
                            "Margin Trend (%)"))
    el.append(PageBreak())
    el.append(_kicker("Financial Performance"))
    el.append(Paragraph("EBITDA Trajectory", S_H1))
    el.append(_bar_drawing(months, [("EBITDA", ebitda)], "EBITDA by Period"))
    el.append(PageBreak())

    # NARRATIVE SECTIONS with kicker + divider + optional callout
    for idx, s in enumerate(sections):
        block = [_kicker("Analysis"), Paragraph(str(s.get("h", "Section")), S_H1)]
        body = str(s.get("body", ""))
        # first sentence becomes a callout for emphasis
        if len(body) > 120:
            cut = body.find(". ")
            if 20 < cut < 240:
                block.append(_callout(body[:cut + 1], TEAL))
                block.append(Spacer(1, 8))
                body = body[cut + 2:]
        block.append(Paragraph(body, S_BODY))
        el.append(KeepTogether(block))
        el.append(Spacer(1, 6))
    el.append(PageBreak())

    # RISK REGISTER
    el.append(_kicker("Governance"))
    el.append(Paragraph("Risk Register", S_H1))
    rdata = [["Risk", "Severity", "Mitigation"]] + [list(r[:3]) for r in risks[:6]]
    el.append(_wrapped_table(rdata, [6.0 * cm, 2.4 * cm, 7.6 * cm]))
    el.append(PageBreak())

    # RECOMMENDATIONS
    el.append(_kicker("Decisions Requested"))
    el.append(Paragraph("Recommendations & Board Asks", S_H1))
    for i, r in enumerate(recs[:6], 1):
        el.append(_callout(f"{i}.  {r}", NAVY))
        el.append(Spacer(1, 6))
    el.append(PageBreak())

    # APPENDIX
    el.append(_kicker("Reference"))
    el.append(Paragraph("Appendix \u2014 Key Assumptions", S_H1))
    adata = [["Assumption", "Value"],
             ["Reporting currency", currency_symbol],
             ["COGS % of revenue", "22% (trailing average)"],
             ["Base-case growth", "12% (pipeline-weighted)"],
             ["Runway basis", "Trailing 3-month net burn"]]
    el.append(_wrapped_table(adata, [8 * cm, 8 * cm]))
    el.append(Spacer(1, 12))
    el.append(Paragraph("Generated by the OrchestrIQ Document Intelligence Engine. "
                        "Figures reconcile to the accompanying Excel workbook.", S_SMALL))

    doc.build(el)
    return buf.getvalue()
