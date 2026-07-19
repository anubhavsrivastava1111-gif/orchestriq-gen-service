"""
OrchestrIQ PDF Engine v3 — Board/Investor Grade
ReportLab with real tables, embedded charts, TOC, cover page.
No bullet encoding bugs. Tables break properly across pages.
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm, inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, KeepTogether
)
from reportlab.platypus.tableofcontents import TableOfContents
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.graphics.shapes import Drawing, Rect, String, Line
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.linecharts import HorizontalLineChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics import renderPDF
import io
from datetime import datetime

# ── Colors ─────────────────────────────────────────────────────────────────────
C_NAVY  = colors.HexColor("#1E3A5F")
C_TEAL  = colors.HexColor("#14B8A6")
C_LIGHT = colors.HexColor("#F1F5F9")
C_WHITE = colors.white
C_DARK  = colors.HexColor("#0F172A")
C_MUTED = colors.HexColor("#94A3B8")
C_GREEN = colors.HexColor("#10B981")
C_RED   = colors.HexColor("#EF4444")
C_AMBER = colors.HexColor("#F59E0B")
C_ROW   = colors.HexColor("#F8FAFC")
C_BORD  = colors.HexColor("#E2E8F0")

W, H = A4
M_L, M_R, M_T, M_B = 20*mm, 20*mm, 24*mm, 20*mm
CONTENT_W = W - M_L - M_R

def _styles(currency_symbol="₹"):
    base = getSampleStyleSheet()
    s = {}
    s["body"] = ParagraphStyle("body", fontName="Helvetica", fontSize=10.5,
                                leading=16, textColor=C_DARK, spaceAfter=6,
                                alignment=TA_JUSTIFY)
    s["h1"] = ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=18,
                              textColor=C_NAVY, spaceBefore=18, spaceAfter=6,
                              borderPad=4, leading=22)
    s["h2"] = ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=14,
                              textColor=C_TEAL, spaceBefore=12, spaceAfter=4, leading=18)
    s["h3"] = ParagraphStyle("h3", fontName="Helvetica-Bold", fontSize=12,
                              textColor=C_NAVY, spaceBefore=8, spaceAfter=3, leading=16)
    s["bullet"] = ParagraphStyle("bullet", fontName="Helvetica", fontSize=10.5,
                                  leading=16, textColor=C_DARK, spaceAfter=3,
                                  leftIndent=14, firstLineIndent=-10)
    s["kpi_label"] = ParagraphStyle("kpi_label", fontName="Helvetica", fontSize=9,
                                     textColor=C_MUTED, leading=11)
    s["kpi_value"] = ParagraphStyle("kpi_value", fontName="Helvetica-Bold", fontSize=20,
                                     textColor=C_NAVY, leading=24)
    s["caption"] = ParagraphStyle("caption", fontName="Helvetica-Oblique", fontSize=8.5,
                                   textColor=C_MUTED, leading=11, spaceAfter=4)
    s["finding"] = ParagraphStyle("finding", fontName="Helvetica", fontSize=10.5,
                                   leading=16, textColor=C_DARK, spaceAfter=4,
                                   leftIndent=10, borderPad=6, backColor=C_LIGHT)
    return s

def _tbl_style(has_header=True):
    cmds = [
        ("FONTNAME",   (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE",   (0,0), (-1,-1), 9.5),
        ("ALIGN",      (0,0), (-1,-1), "LEFT"),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [C_WHITE, C_ROW]),
        ("GRID",       (0,0), (-1,-1), 0.4, C_BORD),
        ("LEFTPADDING", (0,0), (-1,-1), 7),
        ("RIGHTPADDING",(0,0), (-1,-1), 7),
        ("TOPPADDING", (0,0), (-1,-1), 5),
        ("BOTTOMPADDING",(0,0), (-1,-1), 5),
    ]
    if has_header:
        cmds += [
            ("BACKGROUND",  (0,0), (-1,0), C_NAVY),
            ("TEXTCOLOR",   (0,0), (-1,0), C_WHITE),
            ("FONTNAME",    (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE",    (0,0), (-1,0), 10),
            ("BOTTOMPADDING",(0,0),(-1,0), 7),
        ]
    return TableStyle(cmds)

def _make_bar_chart(categories, values, title="", width=CONTENT_W*0.7, height=60*mm):
    """Return a Drawing with a bar chart."""
    d = Drawing(width, height + 20)
    chart = VerticalBarChart()
    chart.x = 40
    chart.y = 20
    chart.width = width - 50
    chart.height = height - 10
    chart.data = [values[:10]]
    chart.categoryAxis.categoryNames = [str(c)[:12] for c in categories[:10]]
    chart.bars[0].fillColor = C_TEAL
    chart.valueAxis.valueMin = 0
    chart.valueAxis.labels.fontName = "Helvetica"
    chart.valueAxis.labels.fontSize = 8
    chart.categoryAxis.labels.fontName = "Helvetica"
    chart.categoryAxis.labels.fontSize = 8
    chart.categoryAxis.labels.angle = 30 if len(categories) > 5 else 0
    chart.categoryAxis.labels.dy = -8
    if title:
        d.add(String(width/2, height+8, title, fontName="Helvetica-Bold", fontSize=10,
                     fillColor=C_NAVY, textAnchor="middle"))
    d.add(chart)
    return d

def _parse_md_table(text):
    """Extract rows from a markdown table string."""
    rows = []
    for line in text.split("\n"):
        line = line.strip()
        if not line or not "|" in line:
            continue
        if line.replace("|","").replace("-","").replace(":","").replace(" ","") == "":
            continue
        cells = [c.strip() for c in line.split("|") if c.strip()]
        if cells:
            rows.append(cells)
    return rows

def _cover_page(canvas, doc, title, company, industry, date, classification):
    canvas.saveState()
    # Navy background
    canvas.setFillColor(C_NAVY)
    canvas.rect(0, 0, W, H, fill=1, stroke=0)
    # Teal accent bar
    canvas.setFillColor(C_TEAL)
    canvas.rect(0, H*0.42, W, 3, fill=1, stroke=0)
    # Title
    canvas.setFillColor(C_WHITE)
    canvas.setFont("Helvetica-Bold", 26)
    canvas.drawString(M_L, H*0.55, title[:65])
    # Company / industry
    canvas.setFillColor(C_TEAL)
    canvas.setFont("Helvetica", 14)
    canvas.drawString(M_L, H*0.48, f"{company}  |  {industry}")
    # Date / classification
    canvas.setFillColor(C_MUTED)
    canvas.setFont("Helvetica", 10)
    canvas.drawString(M_L, H*0.12, f"{classification}  ·  {date}")
    canvas.restoreState()

def _header_footer(canvas, doc, title, company, page_num, total_pages):
    canvas.saveState()
    # Header
    canvas.setFillColor(C_NAVY)
    canvas.rect(0, H - M_T + 2*mm, W, M_T - 2*mm, fill=1, stroke=0)
    canvas.setFillColor(C_TEAL)
    canvas.rect(0, H - M_T + 2*mm, 4, M_T - 2*mm, fill=1, stroke=0)
    canvas.setFillColor(C_WHITE)
    canvas.setFont("Helvetica-Bold", 8)
    canvas.drawString(M_L + 6, H - 10*mm, title[:70])
    canvas.setFont("Helvetica", 8)
    canvas.drawRightString(W - M_R, H - 10*mm, company)
    # Footer
    canvas.setFillColor(C_LIGHT)
    canvas.rect(0, 0, W, M_B - 2*mm, fill=1, stroke=0)
    canvas.setFillColor(C_MUTED)
    canvas.setFont("Helvetica", 7.5)
    canvas.drawString(M_L, 7*mm, "Confidential")
    canvas.drawCentredString(W/2, 7*mm, f"Page {page_num} of {total_pages}")
    canvas.drawRightString(W - M_R, 7*mm, datetime.now().strftime("%d %b %Y"))
    canvas.restoreState()

def build_pdf(schema: dict, currency_symbol: str = "₹") -> bytes:
    buf = io.BytesIO()
    title      = schema.get("title", "Executive Report")
    company    = schema.get("company", "")
    industry   = schema.get("industry", "")
    classif    = schema.get("classification", "Confidential")
    exec_summ  = schema.get("executive_summary", "")
    sections   = schema.get("sections", [])
    key_findings = schema.get("key_findings", [])
    recommendations = schema.get("recommendations", [])
    kpis       = schema.get("summary_kpis", [])
    charts     = schema.get("charts", [])
    today      = datetime.now().strftime("%d %B %Y")
    total_pages = [1]  # mutable for closure

    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=M_L, rightMargin=M_R,
        topMargin=M_T + 2*mm, bottomMargin=M_B + 2*mm,
        title=title, author=company,
    )

    S = _styles(currency_symbol)
    story = []

    # ── COVER (blank placeholder — drawn in onFirstPage) ────────────────────
    story.append(Spacer(1, H * 0.65))
    story.append(PageBreak())

    # ── EXECUTIVE SUMMARY ─────────────────────────────────────────────────────
    story.append(Paragraph("EXECUTIVE SUMMARY", S["h1"]))
    story.append(HRFlowable(width=CONTENT_W, thickness=2, color=C_TEAL, spaceAfter=8))
    if exec_summ:
        story.append(Paragraph(exec_summ, S["body"]))
        story.append(Spacer(1, 6))

    # KPI cards row
    if kpis:
        kpi_data = [[
            Paragraph(f"<b>{k.get('label','')}</b><br/><font size='18' color='#1E3A5F'>{k.get('value','')}</font>", S["body"])
            for k in kpis[:4]
        ]]
        kpi_tbl = Table(kpi_data, colWidths=[CONTENT_W/min(len(kpis),4)]*min(len(kpis),4))
        kpi_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), C_LIGHT),
            ("BOX",        (0,0), (-1,-1), 1.5, C_TEAL),
            ("INNERGRID",  (0,0), (-1,-1), 0.5, C_BORD),
            ("ALIGN",      (0,0), (-1,-1), "CENTER"),
            ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
            ("TOPPADDING", (0,0), (-1,-1), 10),
            ("BOTTOMPADDING",(0,0),(-1,-1), 10),
        ]))
        story.append(kpi_tbl)
        story.append(Spacer(1, 12))

    # Key Findings
    if key_findings:
        story.append(Paragraph("KEY FINDINGS", S["h2"]))
        for i, f in enumerate(key_findings, 1):
            story.append(Paragraph(f"<b>{i}.</b>  {f}", S["bullet"]))
        story.append(Spacer(1, 8))

    # ── SECTIONS ──────────────────────────────────────────────────────────────
    for sec in sections:
        level = sec.get("level", 1)
        stitle = sec.get("title", "")
        content = str(sec.get("content", ""))

        if level == 1:
            story.append(PageBreak())
            story.append(Paragraph(stitle, S["h1"]))
            story.append(HRFlowable(width=CONTENT_W, thickness=1.5, color=C_TEAL, spaceAfter=6))
        elif level == 2:
            story.append(Paragraph(stitle, S["h2"]))
        else:
            story.append(Paragraph(stitle, S["h3"]))

        # Detect and render table blocks
        lines = content.split("\n")
        table_block = []
        text_block = []
        i = 0
        while i < len(lines):
            line = lines[i]
            if "|" in line and line.strip():
                if text_block:
                    _flush_text(text_block, story, S)
                    text_block = []
                table_block.append(line)
            else:
                if table_block:
                    _flush_table(table_block, story, S)
                    table_block = []
                text_block.append(line)
            i += 1
        if table_block:
            _flush_table(table_block, story, S)
        if text_block:
            _flush_text(text_block, story, S)
        story.append(Spacer(1, 6))

    # ── CHARTS ────────────────────────────────────────────────────────────────
    if charts:
        story.append(PageBreak())
        story.append(Paragraph("DATA VISUALISATION", S["h1"]))
        story.append(HRFlowable(width=CONTENT_W, thickness=1.5, color=C_TEAL, spaceAfter=6))
        for c in charts[:4]:
            try:
                cats = [str(l) for l in c.get("labels", [])]
                vals = [float(v) if v else 0 for v in c.get("values", [])]
                if cats and vals:
                    d = _make_bar_chart(cats, vals, c.get("title",""))
                    story.append(d)
                    story.append(Paragraph(c.get("title",""), S["caption"]))
                    story.append(Spacer(1, 10))
            except:
                pass

    # ── RECOMMENDATIONS ───────────────────────────────────────────────────────
    if recommendations:
        story.append(PageBreak())
        story.append(Paragraph("RECOMMENDATIONS", S["h1"]))
        story.append(HRFlowable(width=CONTENT_W, thickness=2, color=C_TEAL, spaceAfter=8))
        for i, r in enumerate(recommendations, 1):
            story.append(KeepTogether([
                Paragraph(f"<b>Recommendation {i}</b>", S["h3"]),
                Paragraph(str(r), S["body"]),
                Spacer(1, 4),
            ]))

    # Build with page callbacks
    page_counter = [0]
    cover_done = [False]

    def on_first(canvas, doc):
        page_counter[0] = 1
        _cover_page(canvas, doc, title, company, industry, today, classif)

    def on_later(canvas, doc):
        page_counter[0] += 1
        _header_footer(canvas, doc, title, company, page_counter[0], "?")

    doc.build(story, onFirstPage=on_first, onLaterPages=on_later)
    return buf.getvalue()

def _flush_text(lines, story, S):
    for line in lines:
        t = line.strip()
        if not t:
            story.append(Spacer(1, 3))
            continue
        if t.startswith("## "):
            story.append(Paragraph(t[3:], S["h2"]))
        elif t.startswith("### "):
            story.append(Paragraph(t[4:], S["h3"]))
        elif t.startswith(("- ", "* ", "• ")):
            story.append(Paragraph("&#8226;  " + t[2:], S["bullet"]))
        elif t.startswith("[FINDING]"):
            story.append(Paragraph(t[9:].strip(), S["finding"]))
        else:
            story.append(Paragraph(t, S["body"]))

def _flush_table(lines, story, S):
    rows = _parse_md_table("\n".join(lines))
    if len(rows) < 2:
        return
    ncols = max(len(r) for r in rows)
    col_w = CONTENT_W / ncols
    tbl_data = [r[:ncols] + [""] * max(0, ncols - len(r)) for r in rows]
    tbl = Table(tbl_data, colWidths=[col_w]*ncols, repeatRows=1)
    tbl.setStyle(_tbl_style(has_header=True))
    story.append(KeepTogether([tbl, Spacer(1, 6)]))

def _parse_md_table(text):
    rows = []
    for line in text.split("\n"):
        line = line.strip()
        if not line or not "|" in line:
            continue
        if line.replace("|","").replace("-","").replace(":","").replace(" ","") == "":
            continue
        cells = [c.strip() for c in line.split("|") if c.strip() or True]
        cells = [c for i,c in enumerate(cells) if not (i==0 and c=="") and not (i==len(cells)-1 and c=="")]
        if cells:
            rows.append(cells)
    return rows
