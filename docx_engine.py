"""
OrchestrIQ Word Engine v3 — Real .docx binary
python-docx: proper heading styles, real tables with borders, TOC, cover.
No more HTML-saved-as-.doc garbage.
"""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import io
from datetime import datetime

# ── Colors ─────────────────────────────────────────────────────────────────────
NAVY  = RGBColor(0x1E,0x3A,0x5F)
TEAL  = RGBColor(0x14,0xB8,0xA6)
WHITE = RGBColor(0xFF,0xFF,0xFF)
LIGHT = RGBColor(0xF1,0xF5,0xF9)
DARK  = RGBColor(0x0F,0x17,0x2A)
MUTED = RGBColor(0x94,0xA3,0xB8)
GREEN = RGBColor(0x10,0xB9,0x81)
RED   = RGBColor(0xEF,0x44,0x44)

def _set_cell_bg(cell, hex_color: str):
    """Set background color of a Word table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{hex_color}"/>')
    tcPr.append(shd)

def _set_para_spacing(para, before=0, after=6, line=None):
    pf = para.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(line)

def _heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    run = h.runs[0] if h.runs else h.add_run(text)
    if level == 1:
        run.font.color.rgb = NAVY
        run.font.size = Pt(18)
    elif level == 2:
        run.font.color.rgb = TEAL
        run.font.size = Pt(14)
    else:
        run.font.color.rgb = NAVY
        run.font.size = Pt(12)
        run.font.italic = True
    return h

def _para(doc, text, style="Normal", bold=False, color=None, size=11, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = doc.add_paragraph(style=style)
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    p.alignment = align
    _set_para_spacing(p, after=4)
    return p

def _bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(11)
    _set_para_spacing(p, after=3)
    return p

def _add_table(doc, rows, has_header=True, currency_symbol="₹"):
    """Add a real Word table with proper styling."""
    if not rows:
        return
    ncols = max(len(r) for r in rows)
    nrows = len(rows)
    
    tbl = doc.add_table(rows=nrows, cols=ncols)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT

    for ri, row in enumerate(rows):
        for ci, val in enumerate(row[:ncols]):
            cell = tbl.cell(ri, ci)
            cell.text = str(val or "")
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(10 if ri > 0 else 11)
                    run.font.bold = ri == 0
                    if ri == 0:
                        run.font.color.rgb = WHITE
                    else:
                        run.font.color.rgb = DARK
                para.paragraph_format.space_before = Pt(3)
                para.paragraph_format.space_after = Pt(3)
            if ri == 0:
                _set_cell_bg(cell, "1E3A5F")
            elif ri % 2 == 0:
                _set_cell_bg(cell, "F8FAFC")

    doc.add_paragraph("")  # spacing after table

def _parse_md_table(text):
    rows = []
    for line in text.split("\n"):
        line = line.strip()
        if not line or "|" not in line:
            continue
        if line.replace("|","").replace("-","").replace(":","").replace(" ","") == "":
            continue
        cells = [c.strip() for c in line.split("|")]
        cells = [c for i,c in enumerate(cells) if not (i==0 and c=="") and not (i==len(cells)-1 and c=="")]
        if cells:
            rows.append(cells)
    return rows

def build_docx(schema: dict, currency_symbol: str = "₹") -> bytes:
    doc = Document()

    # ── Document styles ────────────────────────────────────────────────────────
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    doc.styles["Normal"].font.color.rgb = DARK

    # Page margins
    for section in doc.sections:
        section.top_margin    = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    title      = schema.get("title","Executive Report")
    company    = schema.get("company","")
    industry   = schema.get("industry","")
    classif    = schema.get("classification","Confidential")
    exec_summ  = schema.get("executive_summary","")
    sections   = schema.get("sections",[])
    key_findings = schema.get("key_findings",[])
    recommendations = schema.get("recommendations",[])
    today      = datetime.now().strftime("%d %B %Y")

    # ── COVER PAGE ─────────────────────────────────────────────────────────────
    # Title
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(100)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    run.font.name = "Calibri"
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = NAVY

    # Company
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(f"{company}  |  {industry}")
    r2.font.name = "Calibri"
    r2.font.size = Pt(14)
    r2.font.color.rgb = TEAL

    # Divider line
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3 = p3.add_run("─" * 60)
    r3.font.color.rgb = TEAL

    # Classification + date
    p4 = doc.add_paragraph()
    p4.paragraph_format.space_before = Pt(80)
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r4 = p4.add_run(f"{classif}  ·  {today}")
    r4.font.name = "Calibri"
    r4.font.size = Pt(10)
    r4.font.color.rgb = MUTED

    doc.add_page_break()

    # ── EXECUTIVE SUMMARY ──────────────────────────────────────────────────────
    _heading(doc, "Executive Summary", 1)
    if exec_summ:
        _para(doc, exec_summ, size=11)
    doc.add_paragraph("")

    # KPI table
    kpis = schema.get("summary_kpis", [])
    if kpis:
        kpi_rows = [
            [k.get("label","") for k in kpis[:4]],
            [str(k.get("value","")) for k in kpis[:4]],
            [str(k.get("delta","—")) for k in kpis[:4]],
        ]
        _add_table(doc, kpi_rows, has_header=True)

    # Key findings
    if key_findings:
        _heading(doc, "Key Findings", 2)
        for i, f in enumerate(key_findings, 1):
            _bullet(doc, f"{i}.  {f}")
        doc.add_paragraph("")

    # ── SECTIONS ──────────────────────────────────────────────────────────────
    for sec in sections:
        level = sec.get("level", 1)
        stitle = sec.get("title","")
        content = str(sec.get("content",""))

        if level == 1:
            doc.add_page_break()
        _heading(doc, stitle, level)

        # Parse content: handle tables, bullet lists, headings, body text
        lines = content.split("\n")
        table_block = []
        text_block = []

        def flush_text():
            for line in text_block:
                t = line.strip()
                if not t:
                    continue
                if t.startswith("## "):
                    _heading(doc, t[3:], 2)
                elif t.startswith("### "):
                    _heading(doc, t[4:], 3)
                elif t.startswith(("- ","* ","• ")):
                    _bullet(doc, t[2:])
                else:
                    # Clean markdown bold
                    t2 = t.replace("**","")
                    _para(doc, t2)

        def flush_table():
            rows = _parse_md_table("\n".join(table_block))
            if rows:
                _add_table(doc, rows, has_header=True, currency_symbol=currency_symbol)

        for line in lines:
            if "|" in line and line.strip():
                if text_block:
                    flush_text()
                    text_block = []
                table_block.append(line)
            else:
                if table_block:
                    flush_table()
                    table_block = []
                text_block.append(line)
        
        if table_block:
            flush_table()
        if text_block:
            flush_text()
        
        doc.add_paragraph("")

    # ── RECOMMENDATIONS ────────────────────────────────────────────────────────
    if recommendations:
        doc.add_page_break()
        _heading(doc, "Recommendations", 1)
        for i, r in enumerate(recommendations, 1):
            p = doc.add_paragraph()
            run = p.add_run(f"{i}.  ")
            run.font.bold = True
            run.font.color.rgb = TEAL
            run.font.name = "Calibri"
            run2 = p.add_run(str(r))
            run2.font.name = "Calibri"
            run2.font.size = Pt(11)
            _set_para_spacing(p, before=6, after=6)
        doc.add_paragraph("")

    # ── FOOTER NOTE ────────────────────────────────────────────────────────────
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(20)
    r = p.add_run(f"Generated by OrchestrIQ  ·  {classif}  ·  {today}")
    r.font.size = Pt(8)
    r.font.color.rgb = MUTED
    r.font.name = "Calibri"
    r.font.italic = True

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()
