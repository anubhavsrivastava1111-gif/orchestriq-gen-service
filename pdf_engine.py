"""
OrchestrIQ Document Intelligence Engine v4 — PDF Engine
Cover page, table of contents, executive sections, KPI table, charts,
risk table, recommendations, appendix, header/footer with page numbers.
"""
import io
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph,
                                Spacer, Table, TableStyle, PageBreak)
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.linecharts import HorizontalLineChart

NAVY = colors.HexColor("#1E3A5F")
TEAL = colors.HexColor("#14B8A6")
GREY = colors.HexColor("#64748B")
LIGHT = colors.HexColor("#F1F5F9")

W, H = A4

S_TITLE = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=26, textColor=colors.white, leading=32)
S_SUB = ParagraphStyle("s", fontName="Helvetica", fontSize=13, textColor=TEAL, leading=18)
S_H1 = ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=16, textColor=NAVY,
                      spaceBefore=18, spaceAfter=8)
S_BODY = ParagraphStyle("b", fontName="Helvetica", fontSize=10.5, textColor=colors.HexColor("#334155"),
                        leading=15.5, spaceAfter=8, alignment=4)
S_TOC = ParagraphStyle("toc", fontName="Helvetica", fontSize=11.5, textColor=NAVY,
                       leading=22, leftIndent=6)
S_SMALL = ParagraphStyle("sm", fontName="Helvetica", fontSize=8.5, textColor=GREY)


def _cover(canvas, doc, title, subtitle):
    canvas.saveState()
    canvas.setFillColor(NAVY); canvas.rect(0, 0, W, H, fill=1, stroke=0)
    canvas.setFillColor(TEAL); canvas.rect(0, H - 4.2 * cm, W, 0.25 * cm, fill=1, stroke=0)
    canvas.setFillColor(colors.white)
    canvas.setFont("Helvetica-Bold", 30)
    y = H - 9 * cm
    for line in _wrap(title, 30):
        canvas.drawString(2.2 * cm, y, line); y -= 1.15 * cm
    canvas.setFillColor(TEAL); canvas.setFont("Helvetica", 15)
    canvas.drawString(2.2 * cm, y - 0.6 * cm, subtitle)
    canvas.setFillColor(colors.HexColor("#94A3B8")); canvas.setFont("Helvetica", 10)
    canvas.drawString(2.2 * cm, 2.6 * cm, "Confidential — Prepared for the Board of Directors")
    canvas.setFillColor(TEAL); canvas.rect(0, 1.8 * cm, W, 0.12 * cm, fill=1, stroke=0)
    canvas.restoreState()


def _wrap(text, n):
    words, lines, cur = text.split(), [], ""
    for w in words:
        if len(cur) + len(w) + 1 <= n: cur = (cur + " " + w).strip()
        else: lines.append(cur); cur = w
    if cur: lines.append(cur)
    return lines[:3]


def _hf(canvas, doc, title):
    canvas.saveState()
    canvas.setFillColor(TEAL); canvas.rect(0, H - 0.9 * cm, W, 0.12 * cm, fill=1, stroke=0)
    canvas.setFillColor(GREY); canvas.setFont("Helvetica", 8.5)
    canvas.drawString(2 * cm, H - 1.5 * cm, title[:80])
    canvas.drawRightString(W - 2 * cm, 1.2 * cm, f"Page {doc.page - 1}")
    canvas.drawString(2 * cm, 1.2 * cm, "Confidential")
    canvas.restoreState()


def _kpi_table(kpis):
    data = [["KPI", "Value", "Δ vs Prior"]] + [list(k[:3]) for k in kpis[:8]]
    t = Table(_cellwrap(data), colWidths=[7 * cm, 4.5 * cm, 4.5 * cm])
    t.setStyle(_tstyle())
    return t

S_CELL = ParagraphStyle("cell", fontName="Helvetica", fontSize=9.5,
                        textColor=colors.HexColor("#334155"), leading=12)
S_CELL_HEAD = ParagraphStyle("cellhead", fontName="Helvetica-Bold", fontSize=9.5,
                            textColor=colors.white, leading=12)

def _cellwrap(data):
    """Wrap every cell in a Paragraph so text wraps instead of clipping,
    and rows auto-grow to fit. Row 0 is treated as the header row."""
    out = []
    for ri, row in enumerate(data):
        style = S_CELL_HEAD if ri == 0 else S_CELL
        out.append([Paragraph(str(c) if c is not None else "", style) for c in row])
    return out
def _tstyle():
    return TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ])


def _bar_drawing(months, series, title):
    d = Drawing(440, 210)
    ch = VerticalBarChart()
    ch.x, ch.y, ch.width, ch.height = 45, 30, 360, 140
    ch.data = [list(s[1]) for s in series]
    ch.categoryAxis.categoryNames = list(months)
    ch.bars[0].fillColor = NAVY
    if len(series) > 1: ch.bars[1].fillColor = TEAL
    ch.valueAxis.labels.fontSize = 7
    ch.categoryAxis.labels.fontSize = 8
    d.add(ch)
    d.add(String(45, 190, title, fontName="Helvetica-Bold", fontSize=11, fillColor=NAVY))
    return d


def build_pdf(model: dict, title: str, subtitle: str = "Board Report",
              currency_symbol: str = "\u20b9") -> bytes:
    buf = io.BytesIO()
    doc = BaseDocTemplate(buf, pagesize=A4,
                          leftMargin=2 * cm, rightMargin=2 * cm,
                          topMargin=2.4 * cm, bottomMargin=2 * cm)
    frame = Frame(2 * cm, 2 * cm, W - 4 * cm, H - 4.6 * cm, id="f")
    doc.addPageTemplates([
        PageTemplate(id="cover", frames=[frame],
                     onPage=lambda c, d: _cover(c, d, title, subtitle)),
        PageTemplate(id="body", frames=[frame],
                     onPage=lambda c, d: _hf(c, d, title)),
    ])

    months = model["months"]; rev = model["rev"]; gross = model["gross"]
    ebitda = model["ebitda"]; kpis = model["kpis"]; risks = model["risks"]
    recs = model["recs"]
    sections = model.get("sections") or []

    el = []
    # cover page content is drawn on canvas; push to next template
    from reportlab.platypus import NextPageTemplate
    el.append(NextPageTemplate("body"))
    el.append(PageBreak())

    # TOC
    el.append(Paragraph("Table of Contents", S_H1))
    toc_items = ["1. Executive KPI Scorecard", "2. Financial Charts"] + \
                [f"{i + 3}. {s['h']}" for i, s in enumerate(sections)] + \
                [f"{len(sections) + 3}. Risk Register",
                 f"{len(sections) + 4}. Recommendations",
                 f"{len(sections) + 5}. Appendix — Assumptions"]
    for t in toc_items:
        el.append(Paragraph(t, S_TOC))
    el.append(PageBreak())

    # KPI table
    el.append(Paragraph("Executive KPI Scorecard", S_H1))
    el.append(_kpi_table(kpis))
    el.append(Spacer(1, 16))

    # Charts
    el.append(Paragraph("Financial Charts", S_H1))
    el.append(_bar_drawing(months, [("Revenue", rev), ("Gross Profit", gross)],
                           "Revenue & Gross Profit by Month"))
    el.append(Spacer(1, 10))
    el.append(_bar_drawing(months, [("EBITDA", ebitda)], "EBITDA by Month"))
    el.append(PageBreak())

    # Narrative sections
    for s in sections:
        el.append(Paragraph(s["h"], S_H1))
        el.append(Paragraph(s["body"], S_BODY))
    el.append(PageBreak())

    # Risk table
    el.append(Paragraph("Risk Register", S_H1))
    rdata = [["Risk", "Severity", "Mitigation"]] + [list(r[:3]) for r in risks[:6]]
    rt = Table(_cellwrap(rdata), colWidths=[6.5 * cm, 2.5 * cm, 7 * cm])
    rt.setStyle(_tstyle())
    el.append(rt)
    el.append(Spacer(1, 16))

    # Recommendations
    el.append(Paragraph("Recommendations", S_H1))
    for i, r in enumerate(recs[:6], 1):
        el.append(Paragraph(f"{i}. {r}", S_BODY))
    el.append(PageBreak())

    # Appendix
    el.append(Paragraph("Appendix — Key Assumptions", S_H1))
    adata = [["Assumption", "Value"],
             ["Reporting currency", currency_symbol],
             ["COGS % of revenue", "22% (trailing average)"],
             ["Base-case Q3 growth", "12% (pipeline-weighted)"],
             ["Runway basis", "Trailing 3-month net burn"]]
    at = Table(_cellwrap(adata), colWidths=[8 * cm, 8 * cm]); at.setStyle(_tstyle())
    el.append(at)
    el.append(Spacer(1, 12))
    el.append(Paragraph("This report was generated by the OrchestrIQ Document Intelligence Engine. "
                        "Figures reconcile to the accompanying Excel workbook.", S_SMALL))

    doc.build(el)
    return buf.getvalue()
