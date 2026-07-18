"""
OrchestrIQ PDF Engine — Big4/McKinsey-Grade Report Generator
Architecture: Cover drawn via canvas callback (no Flowable sizing bug),
body content via ReportLab Platypus.
"""

import io
from datetime import datetime
from functools import partial
from typing import Any, Dict, List

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame,
    Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether, NextPageTemplate
)
from reportlab.platypus.flowables import Flowable
from reportlab.lib.colors import HexColor

# ─── BRAND COLORS ─────────────────────────────────────────────────────────────
NAVY      = HexColor("#0D1B2A")
NAVY_MID  = HexColor("#1A2744")
SLATE     = HexColor("#2C3E50")
GOLD      = HexColor("#C9A84C")
GOLD_LT   = HexColor("#F0D080")
WHITE     = HexColor("#FFFFFF")
LIGHT_GR  = HexColor("#F7F8FA")
MID_GR    = HexColor("#D6D8DB")
DARK_GR   = HexColor("#3D3D3D")
GREEN_OK  = HexColor("#1A7A4A")
RED_NG    = HexColor("#C0392B")
AMBER     = HexColor("#D68910")
BLUE_AC   = HexColor("#1F4E79")
TEAL_AC   = HexColor("#0D6E8A")

PAGE_W, PAGE_H = A4
MARGIN_L = MARGIN_R = 1.8 * cm
MARGIN_T = 2.4 * cm
MARGIN_B = 2.2 * cm
BODY_W   = PAGE_W - MARGIN_L - MARGIN_R

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

def _is_num_str(s: str) -> bool:
    s = str(s).strip().lstrip("₹$£€").replace(",","").replace("%","") \
               .replace("(","").replace(")","").replace("x","")
    try: float(s); return True
    except: return False

# ─── STYLES ───────────────────────────────────────────────────────────────────

def _build_styles():
    return {
        "h1": ParagraphStyle("H1", fontSize=12, textColor=WHITE,
            fontName="Helvetica-Bold", leading=17, spaceBefore=10, spaceAfter=4),
        "h2": ParagraphStyle("H2", fontSize=10.5, textColor=GOLD,
            fontName="Helvetica-Bold", leading=15, spaceBefore=9, spaceAfter=3),
        "h3": ParagraphStyle("H3", fontSize=9.5, textColor=NAVY,
            fontName="Helvetica-Bold", leading=14, spaceBefore=7, spaceAfter=2),
        "body": ParagraphStyle("Body", fontSize=9, textColor=DARK_GR,
            fontName="Helvetica", alignment=TA_JUSTIFY, leading=13.5,
            spaceBefore=3, spaceAfter=3),
        "bullet": ParagraphStyle("Bullet", fontSize=9, textColor=DARK_GR,
            fontName="Helvetica", leading=13.5, leftIndent=10,
            spaceBefore=2, spaceAfter=2),
        "table_hdr": ParagraphStyle("TH", fontSize=8, textColor=WHITE,
            fontName="Helvetica-Bold", alignment=TA_CENTER),
        "table_cell": ParagraphStyle("TC", fontSize=8, textColor=DARK_GR,
            fontName="Helvetica", alignment=TA_LEFT),
        "table_num": ParagraphStyle("TN", fontSize=8, textColor=DARK_GR,
            fontName="Helvetica", alignment=TA_RIGHT),
        "caption": ParagraphStyle("Cap", fontSize=7.5, textColor=SLATE,
            fontName="Helvetica-Oblique", spaceBefore=2),
        "kpi_label": ParagraphStyle("KL", fontSize=6.5, textColor=GOLD_LT,
            fontName="Helvetica-Bold", alignment=TA_CENTER),
        "kpi_value": ParagraphStyle("KV", fontSize=15, textColor=WHITE,
            fontName="Helvetica-Bold", alignment=TA_CENTER),
        "footer": ParagraphStyle("Ftr", fontSize=7, textColor=MID_GR,
            fontName="Helvetica", alignment=TA_CENTER),
        "callout": ParagraphStyle("CO", fontSize=9, textColor=NAVY,
            fontName="Helvetica-Bold", leading=14, leftIndent=8),
        "finding_num": ParagraphStyle("FN", fontSize=20, textColor=GOLD,
            fontName="Helvetica-Bold", alignment=TA_CENTER),
    }

# ─── CUSTOM FLOWABLES ─────────────────────────────────────────────────────────

class NavyBanner(Flowable):
    def __init__(self, text, width, height=0.72*cm, bg=NAVY,
                 text_color=WHITE, font_size=10, bold=True, stripe=True):
        super().__init__()
        self.text = text; self.bw = width; self.bh = height
        self.bg = bg; self.text_color = text_color
        self.font_size = font_size; self.bold = bold; self.stripe = stripe
        self.width = width; self.height = height

    def draw(self):
        c = self.canv
        c.setFillColor(self.bg)
        c.rect(0, 0, self.bw, self.bh, fill=1, stroke=0)
        if self.stripe:
            c.setFillColor(GOLD)
            c.rect(0, 0, 3.5*mm, self.bh, fill=1, stroke=0)
        c.setFillColor(self.text_color)
        fname = "Helvetica-Bold" if self.bold else "Helvetica"
        c.setFont(fname, self.font_size)
        c.drawString(6*mm if self.stripe else 3*mm, self.bh * 0.27, self.text)


class GoldDivider(Flowable):
    def __init__(self, width=BODY_W, thickness=0.5*mm):
        super().__init__()
        self.bw = width; self.tk = thickness
        self.width = width; self.height = self.tk + 1.5*mm

    def draw(self):
        self.canv.setStrokeColor(GOLD)
        self.canv.setLineWidth(self.tk)
        self.canv.line(0, self.tk / 2, self.bw, self.tk / 2)


class KPIRow(Flowable):
    def __init__(self, kpis, width=BODY_W, sym="₹"):
        super().__init__()
        self.kpis = kpis[:6]; self.bw = width; self.sym = sym
        n = max(len(kpis), 1)
        self.n = min(n, 6)
        self.height = 2.4 * cm
        self.width  = width

    def draw(self):
        c   = self.canv
        n   = self.n
        gap = 3 * mm
        cw  = (self.bw - (n - 1) * gap) / n
        status_col = {"good": GREEN_OK, "bad": RED_NG,
                      "warning": AMBER,  "neutral": NAVY_MID}

        for i, kpi in enumerate(self.kpis):
            x   = i * (cw + gap)
            h   = self.height
            acc = status_col.get(kpi.get("status", "neutral"), NAVY_MID)

            # Card
            c.setFillColor(NAVY)
            c.roundRect(x, 0, cw, h, 2 * mm, fill=1, stroke=0)
            # Accent top stripe
            c.setFillColor(acc)
            c.rect(x, h - 4 * mm, cw, 4 * mm, fill=1, stroke=0)
            c.roundRect(x, h - 4 * mm, cw, 5 * mm, 2 * mm, fill=1, stroke=0)

            cx = x + cw / 2
            # Label
            c.setFillColor(GOLD_LT)
            c.setFont("Helvetica-Bold", 6)
            label = kpi.get("label", "")[:24].upper()
            c.drawCentredString(cx, h - 10 * mm, label)

            # Value
            c.setFillColor(WHITE)
            c.setFont("Helvetica-Bold", 13)
            val = str(kpi.get("value", "—"))[:14]
            c.drawCentredString(cx, h / 2 - 1 * mm, val)

            # Change
            chg = str(kpi.get("change", ""))
            col = GREEN_OK if "+" in chg else (RED_NG if chg.startswith("-") else GOLD_LT)
            c.setFillColor(col)
            c.setFont("Helvetica-Bold", 7)
            c.drawCentredString(cx, 4.5 * mm, chg[:20])


class FindingBox(Flowable):
    """Numbered finding or recommendation callout box."""
    def __init__(self, number, title, body, width=BODY_W, is_rec=False):
        super().__init__()
        self.number = number; self.title = title; self.body = body
        self.bw = width; self.is_rec = is_rec
        self.width = width; self.height = 2.0 * cm

    def draw(self):
        c  = self.canv
        bg = BLUE_AC if self.is_rec else LIGHT_GR
        bd = GOLD    if self.is_rec else NAVY_MID

        # Outer box
        c.setFillColor(bg)
        c.setStrokeColor(bd)
        c.setLineWidth(0.6)
        c.roundRect(0, 0, self.bw, self.height, 2 * mm, fill=1, stroke=1)

        # Number badge
        badge_w = 10 * mm
        c.setFillColor(NAVY if self.is_rec else NAVY)
        c.rect(0, 0, badge_w, self.height, fill=1, stroke=0)
        c.setFillColor(GOLD)
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(badge_w / 2, self.height / 2 - 4 * mm, str(self.number))

        # Title
        c.setFillColor(WHITE if self.is_rec else NAVY)
        c.setFont("Helvetica-Bold", 8.5)
        c.drawString(badge_w + 4 * mm, self.height - 8 * mm, str(self.title)[:80])

        # Body
        c.setFillColor(DARK_GR if not self.is_rec else MID_GR)
        c.setFont("Helvetica", 7.5)
        body_text = str(self.body)[:160]
        c.drawString(badge_w + 4 * mm, self.height - 15 * mm, body_text[:80])
        if len(body_text) > 80:
            c.drawString(badge_w + 4 * mm, self.height - 21 * mm, body_text[80:160])


# ─── TABLE BUILDER ────────────────────────────────────────────────────────────

def _build_table(headers, rows, col_widths=None, zebra=True,
                 highlight_last=False, highlight_first_col=True) -> Table:
    st = _build_styles()
    data = [[Paragraph(str(h), st["table_hdr"]) for h in headers]]
    for row in rows:
        data.append([
            Paragraph(str(c),
                      st["table_num"] if _is_num_str(str(c)) else st["table_cell"])
            for c in row
        ])

    if col_widths is None:
        n = len(headers)
        fw = BODY_W * 0.32
        rw = (BODY_W - fw) / max(1, n - 1)
        col_widths = [fw] + [rw] * (n - 1)

    tbl = Table(data, colWidths=col_widths, repeatRows=1)
    cmds = [
        ("BACKGROUND",    (0, 0),  (-1, 0),  NAVY),
        ("TEXTCOLOR",     (0, 0),  (-1, 0),  WHITE),
        ("FONTNAME",      (0, 0),  (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0),  (-1, -1), 8),
        ("ALIGN",         (1, 1),  (-1, -1), "RIGHT"),
        ("ALIGN",         (0, 0),  (0, -1),  "LEFT"),
        ("ALIGN",         (0, 0),  (-1, 0),  "CENTER"),
        ("ROWBACKGROUNDS",(0, 1),  (-1, -1), [LIGHT_GR, WHITE] if zebra else [WHITE]),
        ("LINEBELOW",     (0, 0),  (-1, 0),  1.2, GOLD),
        ("LINEBELOW",     (0, 1),  (-1, -1), 0.3, MID_GR),
        ("TOPPADDING",    (0, 0),  (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0),  (-1, -1), 4),
        ("LEFTPADDING",   (0, 0),  (-1, -1), 5),
        ("RIGHTPADDING",  (0, 0),  (-1, -1), 5),
        ("VALIGN",        (0, 0),  (-1, -1), "MIDDLE"),
    ]
    if highlight_last and len(data) > 1:
        nr = len(data) - 1
        cmds += [
            ("BACKGROUND", (0, nr), (-1, nr), SLATE),
            ("TEXTCOLOR",  (0, nr), (-1, nr), WHITE),
            ("FONTNAME",   (0, nr), (-1, nr), "Helvetica-Bold"),
            ("LINEABOVE",  (0, nr), (-1, nr), 1.0, GOLD),
        ]
    if highlight_first_col and len(data) > 1:
        cmds += [("FONTNAME", (0, 1), (0, -2), "Helvetica-Bold"),
                 ("TEXTCOLOR", (0, 1), (0, -2), NAVY)]

    tbl.setStyle(TableStyle(cmds))
    return tbl

# ─── CANVAS CALLBACKS — called on every page by ReportLab ─────────────────────

def _draw_cover_page(canvas, doc, params):
    """Draw full cover page directly on canvas — no Flowable sizing constraints."""
    canvas.saveState()
    w, h = PAGE_W, PAGE_H
    sym  = params.get("currency_symbol", "₹")

    # ── Full navy background ───────────────────────────────────────────────────
    canvas.setFillColor(NAVY)
    canvas.rect(0, 0, w, h, fill=1, stroke=0)

    # ── Gold left stripe ──────────────────────────────────────────────────────
    canvas.setFillColor(GOLD)
    canvas.rect(0, 0, 6 * mm, h, fill=1, stroke=0)

    # ── Dark bottom band ──────────────────────────────────────────────────────
    canvas.setFillColor(NAVY_MID)
    canvas.rect(0, 0, w, 3.2 * cm, fill=1, stroke=0)
    canvas.setStrokeColor(GOLD)
    canvas.setLineWidth(1.0)
    canvas.line(0, 3.2 * cm, w, 3.2 * cm)

    # ── Top meta row ──────────────────────────────────────────────────────────
    canvas.setFillColor(GOLD_LT)
    canvas.setFont("Helvetica-Bold", 9.5)
    company = params.get("company_name", "Company").upper()
    canvas.drawString(12 * mm, h - 16 * mm, company)

    canvas.setFillColor(GOLD)
    canvas.setFont("Helvetica", 8)
    clsf = params.get("classification", "CONFIDENTIAL")
    canvas.drawRightString(w - 8 * mm, h - 16 * mm, clsf)

    # Thin gold line under top area
    canvas.setStrokeColor(GOLD_LT)
    canvas.setLineWidth(0.4)
    canvas.line(12 * mm, h - 22 * mm, w - 8 * mm, h - 22 * mm)

    # ── Document type badge ───────────────────────────────────────────────────
    doc_type = params.get("document_type", "STRATEGIC REPORT").upper()
    canvas.setFillColor(GOLD)
    bx, by = 12 * mm, h * 0.66
    canvas.setFont("Helvetica-Bold", 7)
    badge_w = canvas.stringWidth(doc_type, "Helvetica-Bold", 7) + 8 * mm
    canvas.roundRect(bx, by, badge_w, 5 * mm, 1.5 * mm, fill=1, stroke=0)
    canvas.setFillColor(NAVY)
    canvas.drawString(bx + 4 * mm, by + 1.5 * mm, doc_type)

    # ── Main title (word-wrapped) ─────────────────────────────────────────────
    title  = params.get("title", "Strategic Report")
    title_size = 28
    canvas.setFillColor(WHITE)
    canvas.setFont("Helvetica-Bold", title_size)
    max_w = w - 22 * mm
    words = title.split(); lines = []; cur = ""
    for word in words:
        test = (cur + " " + word).strip()
        if canvas.stringWidth(test, "Helvetica-Bold", title_size) < max_w:
            cur = test
        else:
            if cur: lines.append(cur)
            cur = word
    if cur: lines.append(cur)
    y_title = h * 0.60
    for ln in lines:
        canvas.drawString(12 * mm, y_title, ln)
        y_title -= title_size * 1.3

    # ── Gold divider line ─────────────────────────────────────────────────────
    div_y = y_title - 6 * mm
    canvas.setStrokeColor(GOLD)
    canvas.setLineWidth(1.8)
    canvas.line(12 * mm, div_y, w - 8 * mm, div_y)

    # ── Subtitle ──────────────────────────────────────────────────────────────
    subtitle = params.get("subtitle", "")
    if subtitle:
        canvas.setFillColor(GOLD_LT)
        canvas.setFont("Helvetica-Oblique", 12)
        canvas.drawString(12 * mm, div_y - 8 * mm, subtitle[:80])

    # ── Meta row ──────────────────────────────────────────────────────────────
    date_str = params.get("date", datetime.now().strftime("%B %Y"))
    audience = params.get("audience", "Executive Management")
    canvas.setFillColor(MID_GR)
    canvas.setFont("Helvetica", 8.5)
    meta = f"Prepared for: {audience}   ·   {date_str}"
    canvas.drawString(12 * mm, div_y - 18 * mm, meta)

    # ── Summary KPIs (if available) ───────────────────────────────────────────
    cover_kpis = params.get("cover_kpis", [])
    if cover_kpis:
        kpi_y = h * 0.24
        n     = min(len(cover_kpis), 4)
        kw    = (w - 20 * mm) / n
        for i, kpi in enumerate(cover_kpis[:n]):
            kx = 10 * mm + i * kw
            canvas.setFillColor(NAVY_MID)
            canvas.roundRect(kx, kpi_y - 18 * mm, kw - 4 * mm, 18 * mm,
                             2 * mm, fill=1, stroke=0)
            canvas.setFillColor(GOLD_LT)
            canvas.setFont("Helvetica-Bold", 6)
            lbl = str(kpi.get("label", "")).upper()[:20]
            canvas.drawCentredString(kx + (kw - 4 * mm) / 2,
                                     kpi_y - 7 * mm, lbl)
            canvas.setFillColor(WHITE)
            canvas.setFont("Helvetica-Bold", 11)
            val = str(kpi.get("value", "—"))[:12]
            canvas.drawCentredString(kx + (kw - 4 * mm) / 2,
                                     kpi_y - 14 * mm, val)

    # ── Bottom bar ────────────────────────────────────────────────────────────
    canvas.setFillColor(GOLD_LT)
    canvas.setFont("Helvetica", 7)
    footer_txt = (f"Generated by OrchestrIQ | GorakhAI   ·   "
                  f"{datetime.now().strftime('%d %b %Y %H:%M')}")
    canvas.drawString(12 * mm, 12 * mm, footer_txt)
    canvas.setFillColor(GOLD)
    canvas.setFont("Helvetica-Bold", 7)
    canvas.drawRightString(w - 8 * mm, 12 * mm, "Page 1 of N")
    canvas.restoreState()


def _draw_body_page(canvas, doc, params):
    """Running header + footer for all body pages (page 2+)."""
    canvas.saveState()
    company = params.get("company_name", "")
    title   = params.get("title", "")
    clsf    = params.get("classification", "CONFIDENTIAL")
    date_s  = params.get("date", datetime.now().strftime("%b %Y"))

    # Header bar
    canvas.setFillColor(NAVY)
    hx = MARGIN_L - 4 * mm
    canvas.rect(hx, PAGE_H - MARGIN_T + 3 * mm,
                BODY_W + 8 * mm, 7.5 * mm, fill=1, stroke=0)
    # Gold left accent in header
    canvas.setFillColor(GOLD)
    canvas.rect(hx, PAGE_H - MARGIN_T + 3 * mm, 2.5 * mm, 7.5 * mm, fill=1, stroke=0)

    canvas.setFillColor(WHITE)
    canvas.setFont("Helvetica-Bold", 7)
    canvas.drawString(MARGIN_L, PAGE_H - MARGIN_T + 5 * mm, company.upper())
    canvas.setFont("Helvetica", 7)
    t_short = title[:55] + ("…" if len(title) > 55 else "")
    canvas.drawCentredString(PAGE_W / 2, PAGE_H - MARGIN_T + 5 * mm, t_short)
    canvas.drawRightString(PAGE_W - MARGIN_R, PAGE_H - MARGIN_T + 5 * mm, clsf)

    # Gold underline
    canvas.setStrokeColor(GOLD)
    canvas.setLineWidth(0.6)
    canvas.line(hx, PAGE_H - MARGIN_T + 3 * mm,
                hx + BODY_W + 8 * mm, PAGE_H - MARGIN_T + 3 * mm)

    # Footer
    canvas.setStrokeColor(GOLD)
    canvas.setLineWidth(0.4)
    canvas.line(MARGIN_L, MARGIN_B - 7 * mm,
                PAGE_W - MARGIN_R, MARGIN_B - 7 * mm)

    canvas.setFillColor(SLATE)
    canvas.setFont("Helvetica", 6.5)
    canvas.drawString(MARGIN_L, MARGIN_B - 13 * mm,
                      f"© {datetime.now().year} {company}  ·  {clsf}")
    canvas.drawCentredString(PAGE_W / 2, MARGIN_B - 13 * mm,
                             f"OrchestrIQ | GorakhAI   ·   {date_s}")
    canvas.setFont("Helvetica-Bold", 6.5)
    canvas.setFillColor(NAVY)
    canvas.drawRightString(PAGE_W - MARGIN_R, MARGIN_B - 13 * mm,
                           f"Page {doc.page}")
    canvas.restoreState()


# ─── SECTION BUILDERS ─────────────────────────────────────────────────────────

def _section_banner(text, level=1, width=BODY_W):
    configs = {
        1: dict(height=0.78 * cm, bg=NAVY,     font_size=10, bold=True),
        2: dict(height=0.63 * cm, bg=SLATE,    font_size=9,  bold=True),
        3: dict(height=0.55 * cm, bg=NAVY_MID, font_size=8.5,bold=False),
    }
    cfg = configs.get(level, configs[1])
    return NavyBanner(text.upper() if level == 1 else text, width, **cfg)


def _parse_content(content: str, st: dict) -> list:
    """Convert text/markdown-light content to ReportLab elements."""
    elements = []
    for line in content.split("\n"):
        line = line.rstrip()
        if not line:
            elements.append(Spacer(1, 2.5 * mm))
        elif line.startswith("## "):
            elements.append(Paragraph(line[3:], st["h2"]))
        elif line.startswith("# "):
            elements.append(Paragraph(line[2:], st["h3"]))
        elif line.startswith(("- ", "• ", "* ")):
            elements.append(Paragraph(f"• {line[2:]}", st["bullet"]))
        elif line.startswith("**") and line.endswith("**"):
            elements.append(Paragraph(f"<b>{line[2:-2]}</b>", st["h3"]))
        else:
            text = line.replace("**", "<b>", 1).replace("**", "</b>", 1) \
                       .replace("*", "<i>", 1).replace("*", "</i>", 1)
            elements.append(Paragraph(text, st["body"]))
    return elements


# ─── MAIN BUILDER ─────────────────────────────────────────────────────────────

def build_pdf(params: Dict[str, Any]) -> bytes:
    buf = io.BytesIO()
    sym = params.get("currency_symbol", "₹")
    st  = _build_styles()

    # ── Two page templates: cover (no header/footer) and body ─────────────────
    cover_frame = Frame(MARGIN_L, MARGIN_B, BODY_W,
                        PAGE_H - MARGIN_T - MARGIN_B, id="cover")
    body_frame  = Frame(MARGIN_L, MARGIN_B, BODY_W,
                        PAGE_H - MARGIN_T - MARGIN_B, id="body")

    cover_tmpl = PageTemplate(
        id="Cover", frames=[cover_frame],
        onPage=partial(_draw_cover_page, params=params)
    )
    body_tmpl = PageTemplate(
        id="Body", frames=[body_frame],
        onPage=partial(_draw_body_page, params=params)
    )

    doc = BaseDocTemplate(
        buf, pagesize=A4,
        leftMargin=MARGIN_L, rightMargin=MARGIN_R,
        topMargin=MARGIN_T,  bottomMargin=MARGIN_B,
        title=params.get("title", "Report"),
        author="OrchestrIQ | GorakhAI",
    )
    doc.addPageTemplates([cover_tmpl, body_tmpl])

    # ── Flowable list ─────────────────────────────────────────────────────────
    elements = []

    # Page 1 = Cover (drawn entirely via canvas callback; just need a tiny token
    # flowable so ReportLab processes the page, then switch to Body template)
    elements.append(Spacer(1, 1))
    elements.append(NextPageTemplate("Body"))
    elements.append(PageBreak())

    # ── Executive Summary ─────────────────────────────────────────────────────
    elements.append(_section_banner("EXECUTIVE SUMMARY", 1))
    elements.append(Spacer(1, 3 * mm))

    exec_summary = params.get("executive_summary", "")
    if exec_summary:
        for el in _parse_content(exec_summary, st):
            elements.append(el)
        elements.append(Spacer(1, 4 * mm))

    # ── KPI Dashboard Row ─────────────────────────────────────────────────────
    fd   = params.get("financial_data", {})
    kpis = fd.get("kpis", [])
    if kpis:
        elements.append(Spacer(1, 2 * mm))
        elements.append(KPIRow(kpis[:6], BODY_W, sym))
        elements.append(Spacer(1, 5 * mm))

    # ── Key Findings ──────────────────────────────────────────────────────────
    findings = params.get("key_findings", [])
    if findings:
        elements.append(NavyBanner("KEY FINDINGS", BODY_W, 0.6 * cm,
                                   bg=NAVY_MID, font_size=9))
        elements.append(Spacer(1, 3 * mm))
        for i, f in enumerate(findings, 1):
            if isinstance(f, dict):
                title_f = f.get("title", f"Finding {i}")
                body_f  = f.get("body", "")
            else:
                title_f = f"Finding {i}"
                body_f  = str(f)
            box = FindingBox(i, title_f, body_f, BODY_W, is_rec=False)
            elements.append(KeepTogether([box, Spacer(1, 2.5 * mm)]))
        elements.append(Spacer(1, 3 * mm))

    # ── Financial Snapshot Table ───────────────────────────────────────────────
    rev = fd.get("revenue", [])
    if rev:
        gp     = fd.get("gross_profit",  [0] * len(rev))
        ebitda = fd.get("ebitda",        [0] * len(rev))
        net    = fd.get("net_profit",    [0] * len(rev))
        gm_pct = fd.get("gp_margin",     [0] * len(rev))
        labels = fd.get("period_labels", [f"P{i+1}" for i in range(len(rev))])

        elements.append(_section_banner("FINANCIAL SNAPSHOT", 2))
        elements.append(Spacer(1, 2 * mm))

        n_periods = min(len(rev), 6)
        headers   = ["Metric"] + labels[:n_periods]
        fw = BODY_W * 0.28
        cw = (BODY_W - fw) / max(1, n_periods)

        rows = [
            ["Revenue"]             + [_fmt(_num(v), sym) for v in rev[:n_periods]],
            ["  Gross Profit"]      + [_fmt(_num(v), sym) for v in gp[:n_periods]],
            ["  GP Margin %"]       + [f"{_num(v):.1f}%" for v in gm_pct[:n_periods]],
            ["EBITDA"]              + [_fmt(_num(v), sym) for v in ebitda[:n_periods]],
            ["Net Profit / (Loss)"] + [_fmt(_num(v), sym) for v in net[:n_periods]],
        ]
        tbl = _build_table(headers, rows, [fw] + [cw] * n_periods,
                           highlight_last=True)
        elements.append(tbl)
        elements.append(Spacer(1, 4 * mm))

    # ── Margin table if available ──────────────────────────────────────────────
    margins = fd.get("margins", {})
    if margins and rev:
        ebitda_m = fd.get("ebitda_margin", [0] * len(rev))
        net_m    = fd.get("net_margin",    [0] * len(rev))
        labels   = fd.get("period_labels", [f"P{i+1}" for i in range(len(rev))])
        n_p      = min(len(rev), 6)
        headers  = ["Margin Analysis"] + labels[:n_p]
        fw2 = BODY_W * 0.28; cw2 = (BODY_W - fw2) / max(1, n_p)
        rows2 = [
            ["EBITDA Margin %"] + [f"{_num(v):.1f}%" for v in ebitda_m[:n_p]],
            ["Net Profit Margin %"] + [f"{_num(v):.1f}%" for v in net_m[:n_p]],
        ]
        tbl2 = _build_table(headers, rows2, [fw2] + [cw2] * n_p)
        elements.append(tbl2)
        elements.append(Spacer(1, 3 * mm))

    elements.append(GoldDivider(BODY_W))
    elements.append(Spacer(1, 4 * mm))

    # ── Main Sections ─────────────────────────────────────────────────────────
    sections = params.get("sections", [])
    for sec in sections:
        level    = sec.get("level", 1)
        s_title  = sec.get("title", "")
        content  = sec.get("content", "")
        with_sec = []

        with_sec.append(_section_banner(s_title, level))
        with_sec.append(Spacer(1, 3 * mm))

        if content:
            with_sec.extend(_parse_content(content, st))
            with_sec.append(Spacer(1, 3 * mm))

        # Section KPIs
        sec_kpis = sec.get("kpis", [])
        if sec_kpis:
            with_sec.append(KPIRow(sec_kpis[:6], BODY_W, sym))
            with_sec.append(Spacer(1, 4 * mm))

        # Section tables
        for tbl_data in sec.get("tables", []):
            if tbl_data.get("title"):
                with_sec.append(Paragraph(
                    f"<i>{tbl_data['title']}</i>", st["caption"]))
                with_sec.append(Spacer(1, 1 * mm))
            hdrs = tbl_data.get("headers", [])
            rws  = tbl_data.get("rows", [])
            if hdrs and rws:
                with_sec.append(_build_table(hdrs, rws))
            with_sec.append(Spacer(1, 3 * mm))

        elements.extend(with_sec)

    # ── Recommendations ───────────────────────────────────────────────────────
    recs = params.get("recommendations", [])
    if recs:
        elements.append(PageBreak())
        elements.append(_section_banner("STRATEGIC RECOMMENDATIONS", 1))
        elements.append(Spacer(1, 3 * mm))
        for i, rec in enumerate(recs, 1):
            if isinstance(rec, dict):
                title_r = rec.get("title", f"Recommendation {i}")
                body_r  = rec.get("body", "")
            else:
                title_r = f"Recommendation {i}"
                body_r  = str(rec)
            box = FindingBox(i, title_r, body_r, BODY_W, is_rec=True)
            elements.append(KeepTogether([box, Spacer(1, 3 * mm)]))

    # ── Action Plan Table ──────────────────────────────────────────────────────
    action_plan = params.get("action_plan", [])
    if action_plan:
        elements.append(Spacer(1, 3 * mm))
        elements.append(_section_banner("ACTION PLAN", 2))
        elements.append(Spacer(1, 2 * mm))
        hdrs = ["#", "Action", "Owner", "Timeline", "Priority"]
        fw3 = [0.8 * cm, BODY_W * 0.40, BODY_W * 0.18,
               BODY_W * 0.16, BODY_W * 0.14]
        rows3 = [[str(i+1)] + [a.get(k, "") for k in
                  ["action", "owner", "timeline", "priority"]]
                 for i, a in enumerate(action_plan)]
        elements.append(_build_table(hdrs, rows3, fw3))
        elements.append(Spacer(1, 4 * mm))

    # ── Appendices ────────────────────────────────────────────────────────────
    appendices = params.get("appendices", [])
    if appendices:
        elements.append(PageBreak())
        elements.append(_section_banner("APPENDICES", 1))
        elements.append(Spacer(1, 4 * mm))
        for app in appendices:
            elements.append(NavyBanner(
                app.get("title", "Appendix"), BODY_W, 0.6 * cm,
                bg=NAVY_MID, font_size=9))
            elements.append(Spacer(1, 2 * mm))
            if app.get("content"):
                elements.extend(_parse_content(app["content"], st))
            for tbl_d in app.get("tables", []):
                if tbl_d.get("title"):
                    elements.append(Paragraph(
                        f"<i>{tbl_d['title']}</i>", st["caption"]))
                hdrs = tbl_d.get("headers", [])
                rws  = tbl_d.get("rows", [])
                if hdrs and rws:
                    elements.append(_build_table(hdrs, rws))
            elements.append(Spacer(1, 4 * mm))

    # ── Build ─────────────────────────────────────────────────────────────────
    doc.build(elements)
    buf.seek(0)
    return buf.read()
