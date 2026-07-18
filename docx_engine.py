"""
OrchestrIQ DOCX Engine — McKinsey/Big4-Grade Word Document Builder
python-docx: Full cover page, structured consulting report, tables, callout boxes.
"""

import io
from datetime import datetime
from typing import Any, Dict, List

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Emu, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH as WD_ALIGN
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ─── BRAND COLORS ────────────────────────────────────────────────────────────
def _rgb(r, g, b): return RGBColor(r, g, b)

NAVY     = _rgb(13, 27, 42)
NAVY_MID = _rgb(26, 39, 68)
SLATE    = _rgb(44, 62, 80)
GOLD     = _rgb(201, 168, 76)
GOLD_LT  = _rgb(240, 208, 128)
WHITE    = _rgb(255, 255, 255)
LIGHT_GR = _rgb(247, 248, 250)
MID_GR   = _rgb(214, 216, 219)
DARK_GR  = _rgb(58, 58, 58)
GREEN_POS = _rgb(26, 122, 74)
GREEN_LT  = _rgb(230, 244, 237)
RED_NEG   = _rgb(192, 57, 43)
RED_LT    = _rgb(253, 232, 230)
AMBER     = _rgb(214, 137, 16)
AMBER_LT  = _rgb(254, 249, 231)
BLUE_AC   = _rgb(31, 78, 121)

# ─── HELPERS ─────────────────────────────────────────────────────────────────

def _num(v):
    try: return float(str(v or 0).replace(",","").replace("%",""))
    except: return 0.0

def _fmt(v, sym="₹"):
    v = _num(v)
    if abs(v) >= 1e7:  return f"{sym}{v/1e7:.1f}Cr"
    if abs(v) >= 1e5:  return f"{sym}{v/1e5:.1f}L"
    if abs(v) >= 1000: return f"{sym}{v/1000:.1f}K"
    return f"{sym}{v:.0f}"

def _hex(rgb: RGBColor) -> str:
    return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"

def _set_cell_bg(cell, color: RGBColor):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), _hex(color))
    tcPr.append(shd)

def _set_cell_border(cell, color_hex="C9A84C", sz=4):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top","left","bottom","right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color_hex)
        tcBorders.append(el)
    tcPr.append(tcBorders)

def _para_shading(para, color_hex: str):
    """Apply background color to paragraph via pPr/shd element."""
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    pPr.append(shd)

def _set_col_width(col, width_cm: float):
    for cell in col.cells:
        cell.width = Cm(width_cm)

# ─── DOCUMENT SETUP ──────────────────────────────────────────────────────────

def _setup_doc(params: dict) -> Document:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width    = Cm(21)    # A4 width
    sec.page_height   = Cm(29.7) # A4 height
    sec.left_margin   = Cm(2.0)
    sec.right_margin  = Cm(2.0)
    sec.top_margin    = Cm(2.4)
    sec.bottom_margin = Cm(2.2)

    # Set default styles
    style = doc.styles["Normal"]
    style.font.name  = "Calibri"
    style.font.size  = Pt(10)
    style.font.color.rgb = DARK_GR

    return doc

# ─── COVER PAGE ──────────────────────────────────────────────────────────────

def _add_cover(doc: Document, params: dict) -> None:
    """Full-page navy cover page."""
    company  = params.get("company_name", "Company")
    title    = params.get("title",        "Strategic Report")
    subtitle = params.get("subtitle",     "")
    audience = params.get("audience",     "Executive Management")
    date_str = params.get("date",         datetime.now().strftime("%B %Y"))
    clsf     = params.get("classification","CONFIDENTIAL")
    doc_type = params.get("document_type","STRATEGIC REPORT")

    def cover_para(text, size=11, bold=False, color=WHITE,
                   align=WD_ALIGN.LEFT, space_before=0, space_after=0,
                   italic=False, bg=None):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after  = Pt(space_after)
        if bg: _para_shading(p, bg)
        run = p.add_run(str(text))
        run.bold = bold; run.italic = italic
        run.font.size = Pt(size)
        run.font.color.rgb = color
        run.font.name = "Calibri"
        return p

    # Navy background paragraphs for cover
    cover_para(" ", bg=_hex(NAVY), space_before=0, space_after=0)

    # Classification top-right
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN.RIGHT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    _para_shading(p, _hex(NAVY))
    r = p.add_run(clsf)
    r.font.size = Pt(8); r.font.bold = True
    r.font.color.rgb = GOLD; r.font.name = "Calibri"

    # Company name
    cover_para(company.upper(), size=11, bold=True, color=GOLD_LT,
               bg=_hex(NAVY), space_before=4, space_after=2)

    # Divider
    cover_para("—" * 60, size=7, color=GOLD, bg=_hex(NAVY), space_before=2, space_after=4)

    # Document type badge
    cover_para(f"[ {doc_type} ]", size=8, bold=True, color=GOLD,
               bg=_hex(NAVY), space_before=6, space_after=8)

    # Main title (large)
    cover_para(title, size=28, bold=True, color=WHITE,
               bg=_hex(NAVY), space_before=8, space_after=6)

    # Subtitle
    if subtitle:
        cover_para(subtitle, size=13, italic=True, color=GOLD_LT,
                   bg=_hex(NAVY), space_before=4, space_after=4)

    # Gold separator
    cover_para("━" * 80, size=7, color=GOLD, bg=_hex(NAVY),
               space_before=8, space_after=8)

    # Prepared for
    cover_para(f"Prepared for: {audience}", size=10, color=MID_GR,
               bg=_hex(NAVY_MID), space_before=4, space_after=2)
    cover_para(date_str, size=11, bold=True, color=GOLD_LT,
               bg=_hex(NAVY_MID), space_before=2, space_after=6)

    # Large spacer
    for _ in range(8):
        cover_para(" ", bg=_hex(NAVY_MID), space_before=0, space_after=0)

    # Footer
    cover_para(
        f"Generated by OrchestrIQ | GorakhAI   ·   {datetime.now().strftime('%d %b %Y')}",
        size=7.5, color=GOLD_LT, align=WD_ALIGN.CENTER,
        bg=_hex(NAVY), space_before=8, space_after=2)

    doc.add_page_break()

# ─── SECTION STYLES ──────────────────────────────────────────────────────────

def _h1(doc: Document, text: str) -> None:
    """Level-1 section heading: full navy banner."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN.LEFT
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(4)
    _para_shading(p, _hex(NAVY))
    run = p.add_run(f"  {text.upper()}")
    run.bold = True; run.font.size = Pt(11)
    run.font.color.rgb = WHITE; run.font.name = "Calibri"

def _h2(doc: Document, text: str) -> None:
    """Level-2 heading: slate band."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN.LEFT
    p.paragraph_format.space_before = Pt(9)
    p.paragraph_format.space_after  = Pt(3)
    _para_shading(p, _hex(SLATE))
    run = p.add_run(f"  {text}")
    run.bold = True; run.font.size = Pt(10)
    run.font.color.rgb = GOLD_LT; run.font.name = "Calibri"

def _h3(doc: Document, text: str) -> None:
    """Level-3 heading: gold underline."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN.LEFT
    p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after  = Pt(2)
    run = p.add_run(text)
    run.bold = True; run.font.size = Pt(10)
    run.font.color.rgb = NAVY; run.font.name = "Calibri"
    # Gold bottom border via pPr
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot  = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single")
    bot.set(qn("w:sz"), "6")
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), _hex(GOLD))
    pBdr.append(bot)
    pPr.append(pBdr)

def _body(doc: Document, text: str, indent=False) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN.JUSTIFY
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after  = Pt(3)
    if indent:
        p.paragraph_format.left_indent = Cm(0.5)
    run = p.add_run(str(text))
    run.font.size = Pt(9.5); run.font.color.rgb = DARK_GR
    run.font.name = "Calibri"

def _bullet(doc: Document, text: str, level=1) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    ind = Cm(0.4 * level)
    p.paragraph_format.left_indent  = ind + Cm(0.4)
    p.paragraph_format.first_line_indent = Cm(-0.4)
    marker = "▪" if level == 1 else "○"
    run = p.add_run(f"{marker}  {text}")
    run.font.size = Pt(9.5); run.font.color.rgb = DARK_GR
    run.font.name = "Calibri"

def _divider(doc: Document) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(6)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot  = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single"); bot.set(qn("w:sz"), "6")
    bot.set(qn("w:space"), "1"); bot.set(qn("w:color"), _hex(GOLD))
    pBdr.append(bot); pPr.append(pBdr)

# ─── TABLE BUILDER ───────────────────────────────────────────────────────────

def _add_table(doc: Document, headers: List[str], rows: List[List],
               title: str = "", col_widths: List[float] = None,
               highlight_last: bool = False) -> None:
    if title:
        p = doc.add_paragraph()
        r = p.add_run(f"  {title}")
        r.bold = True; r.font.size = Pt(8.5)
        r.font.color.rgb = NAVY; r.font.name = "Calibri"
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after  = Pt(2)

    n_cols = len(headers)
    n_rows = len(rows) + 1
    tbl    = doc.add_table(rows=n_rows, cols=n_cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.style     = "Table Grid"

    # Auto column widths
    page_w = 17.0  # cm (A4 - margins)
    if col_widths is None:
        first_w = page_w * 0.32
        rest_w  = (page_w - first_w) / max(1, n_cols - 1)
        col_widths = [first_w] + [rest_w] * (n_cols - 1)

    # Header row
    hdr_row = tbl.rows[0]
    for ci, hdr in enumerate(headers):
        cell = hdr_row.cells[ci]
        _set_cell_bg(cell, NAVY)
        _set_cell_border(cell, _hex(NAVY_MID), 4)
        cell.width = Cm(col_widths[ci])
        p2 = cell.paragraphs[0]
        p2.alignment = WD_ALIGN.CENTER
        p2.paragraph_format.space_before = Pt(2)
        p2.paragraph_format.space_after  = Pt(2)
        run = p2.add_run(str(hdr))
        run.bold = True; run.font.size = Pt(8.5)
        run.font.color.rgb = WHITE; run.font.name = "Calibri"

    # Data rows
    for ri, row in enumerate(rows):
        tbl_row = tbl.rows[ri + 1]
        is_last = (ri == len(rows) - 1) and highlight_last
        is_alt  = ri % 2 == 1
        row_bg  = _rgb(208,215,226) if is_last else (LIGHT_GR if is_alt else WHITE)

        for ci, cell_val in enumerate(row):
            cell = tbl_row.cells[ci]
            _set_cell_bg(cell, row_bg)
            _set_cell_border(cell, _hex(MID_GR), 2)
            cell.width = Cm(col_widths[ci])
            p2 = cell.paragraphs[0]
            is_num = False
            try: float(str(cell_val).replace(",","").replace("%","").lstrip("₹$£€")); is_num = True
            except: pass
            p2.alignment = WD_ALIGN.RIGHT if (is_num and ci > 0) else WD_ALIGN.LEFT
            p2.paragraph_format.space_before = Pt(1.5)
            p2.paragraph_format.space_after  = Pt(1.5)
            run = p2.add_run(str(cell_val))
            run.bold = is_last; run.font.size = Pt(8.5)
            run.font.color.rgb = NAVY if is_last else DARK_GR
            run.font.name = "Calibri"

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

# ─── KPI TABLE ───────────────────────────────────────────────────────────────

def _add_kpi_table(doc: Document, kpis: List[dict], sym="₹") -> None:
    """Compact 3-column KPI summary bar."""
    n = min(len(kpis), 6)
    if not n: return
    tbl = doc.add_table(rows=2, cols=n)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    w_each = 17.0 / n

    for i, kpi in enumerate(kpis[:n]):
        # Label row
        lbl_cell = tbl.rows[0].cells[i]
        _set_cell_bg(lbl_cell, NAVY)
        lbl_cell.width = Cm(w_each)
        p = lbl_cell.paragraphs[0]
        p.alignment = WD_ALIGN.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after  = Pt(1)
        r = p.add_run(str(kpi.get("label","")).upper()[:24])
        r.font.size = Pt(7); r.bold = True
        r.font.color.rgb = GOLD_LT; r.font.name = "Calibri"

        # Value row
        val_cell = tbl.rows[1].cells[i]
        _set_cell_bg(val_cell, NAVY_MID)
        val_cell.width = Cm(w_each)
        pv = val_cell.paragraphs[0]
        pv.alignment = WD_ALIGN.CENTER
        pv.paragraph_format.space_before = Pt(2)
        pv.paragraph_format.space_after  = Pt(2)
        rv = pv.add_run(str(kpi.get("value","—")))
        rv.font.size = Pt(13); rv.bold = True
        rv.font.color.rgb = WHITE; rv.font.name = "Calibri"
        # Change
        chg = str(kpi.get("change",""))
        if chg:
            rch = pv.add_run(f"\n{chg}")
            rch.font.size = Pt(7.5); rch.bold = False
            col = GREEN_POS if "+" in chg else (RED_NEG if chg.startswith("-") else GOLD_LT)
            rch.font.color.rgb = col; rch.font.name = "Calibri"

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

# ─── CALLOUT BOXES ───────────────────────────────────────────────────────────

def _finding_box(doc: Document, number: int, title: str, body: str) -> None:
    """Blue-shaded finding callout."""
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    # Badge col
    badge = tbl.rows[0].cells[0]
    badge.width = Cm(1.0)
    _set_cell_bg(badge, NAVY)
    p = badge.paragraphs[0]
    p.alignment = WD_ALIGN.CENTER
    p.paragraph_format.space_before = Pt(4)
    r = p.add_run(str(number))
    r.font.size = Pt(18); r.bold = True
    r.font.color.rgb = GOLD; r.font.name = "Calibri"

    # Content col
    content = tbl.rows[0].cells[1]
    content.width = Cm(16)
    _set_cell_bg(content, _rgb(235,241,250))
    _set_cell_border(content, _hex(BLUE_AC), 4)
    pt = content.paragraphs[0]
    pt.paragraph_format.space_before = Pt(3)
    pt.paragraph_format.space_after  = Pt(1)
    rt = pt.add_run(str(title)[:100])
    rt.bold = True; rt.font.size = Pt(10)
    rt.font.color.rgb = NAVY; rt.font.name = "Calibri"
    pb = content.add_paragraph()
    pb.paragraph_format.space_before = Pt(1)
    pb.paragraph_format.space_after  = Pt(3)
    rb = pb.add_run(str(body)[:300])
    rb.font.size = Pt(9); rb.font.color.rgb = DARK_GR
    rb.font.name = "Calibri"

    doc.add_paragraph().paragraph_format.space_after = Pt(3)

def _rec_box(doc: Document, number: int, title: str, body: str) -> None:
    """Green-shaded recommendation callout."""
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    badge = tbl.rows[0].cells[0]
    badge.width = Cm(1.0)
    _set_cell_bg(badge, GREEN_POS)
    p = badge.paragraphs[0]
    p.alignment = WD_ALIGN.CENTER
    p.paragraph_format.space_before = Pt(4)
    r = p.add_run(str(number))
    r.font.size = Pt(18); r.bold = True
    r.font.color.rgb = WHITE; r.font.name = "Calibri"

    content = tbl.rows[0].cells[1]
    content.width = Cm(16)
    _set_cell_bg(content, _rgb(230,244,237))
    _set_cell_border(content, _hex(GREEN_POS), 4)
    pt = content.paragraphs[0]
    pt.paragraph_format.space_before = Pt(3)
    pt.paragraph_format.space_after  = Pt(1)
    rt = pt.add_run(str(title)[:100])
    rt.bold = True; rt.font.size = Pt(10)
    rt.font.color.rgb = GREEN_POS; rt.font.name = "Calibri"
    pb = content.add_paragraph()
    pb.paragraph_format.space_before = Pt(1)
    pb.paragraph_format.space_after  = Pt(3)
    rb = pb.add_run(str(body)[:300])
    rb.font.size = Pt(9); rb.font.color.rgb = DARK_GR
    rb.font.name = "Calibri"

    doc.add_paragraph().paragraph_format.space_after = Pt(3)

# ─── HEADER / FOOTER ─────────────────────────────────────────────────────────

def _setup_header_footer(doc: Document, params: dict) -> None:
    """Add running header and footer to all body sections."""
    for section in doc.sections:
        section.different_first_page_header_footer = True

        # Header
        header = section.header
        header.is_linked_to_previous = False
        header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        header_para.clear()
        header_para.alignment = WD_ALIGN.LEFT
        _para_shading(header_para, _hex(NAVY))

        r1 = header_para.add_run(f"  {params.get('company_name','').upper()}")
        r1.bold = True; r1.font.size = Pt(7.5)
        r1.font.color.rgb = WHITE; r1.font.name = "Calibri"

        r2 = header_para.add_run(
            f"{'  ' * 10}{params.get('title','')[:50]}  |  "
            f"{params.get('classification','CONFIDENTIAL')}")
        r2.font.size = Pt(7); r2.font.color.rgb = GOLD_LT
        r2.font.name = "Calibri"
        header_para.paragraph_format.space_before = Pt(0)
        header_para.paragraph_format.space_after  = Pt(0)

        # Footer
        footer = section.footer
        footer.is_linked_to_previous = False
        footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        footer_para.clear()
        footer_para.alignment = WD_ALIGN.CENTER
        _para_shading(footer_para, _hex(LIGHT_GR))

        # Add gold top border to footer
        pPr = footer_para._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        top  = OxmlElement("w:top")
        top.set(qn("w:val"), "single"); top.set(qn("w:sz"), "6")
        top.set(qn("w:space"), "1"); top.set(qn("w:color"), _hex(GOLD))
        pBdr.append(top); pPr.append(pBdr)

        rftr = footer_para.add_run(
            f"OrchestrIQ | GorakhAI   ·   "
            f"{params.get('company_name','')}   ·   "
            f"{params.get('date', datetime.now().strftime('%B %Y'))}   ·   Page ")
        rftr.font.size = Pt(7); rftr.font.color.rgb = SLATE
        rftr.font.name = "Calibri"

        # Page number field
        fld = OxmlElement("w:fldChar")
        fld.set(qn("w:fldCharType"), "begin")
        footer_para.runs[-1]._r.append(fld)
        instr = OxmlElement("w:instrText")
        instr.text = "PAGE"
        r_fld = OxmlElement("w:r")
        r_fld.append(instr)
        footer_para._p.append(r_fld)
        fld2 = OxmlElement("w:fldChar")
        fld2.set(qn("w:fldCharType"), "end")
        r_fld2 = OxmlElement("w:r")
        r_fld2.append(fld2)
        footer_para._p.append(r_fld2)

def _parse_content(doc: Document, content: str) -> None:
    """Parse light-markdown content and add to doc."""
    for line in content.split("\n"):
        line = line.rstrip()
        if not line:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(2)
        elif line.startswith("## "):
            _h3(doc, line[3:])
        elif line.startswith("# "):
            _h3(doc, line[2:])
        elif line.startswith(("- ", "• ", "* ")):
            _bullet(doc, line[2:])
        elif line.startswith("**") and line.endswith("**"):
            _h3(doc, line[2:-2])
        else:
            # Handle inline bold
            text = line.replace("**", "", 2)
            _body(doc, text)

# ─── MAIN BUILDER ────────────────────────────────────────────────────────────

def build_docx(params: Dict[str, Any]) -> bytes:
    doc = _setup_doc(params)
    sym = params.get("currency_symbol", "₹")

    # ── Cover Page ────────────────────────────────────────────────────────────
    _add_cover(doc, params)
    _setup_header_footer(doc, params)

    # ── Table of Contents (placeholder) ──────────────────────────────────────
    _h1(doc, "TABLE OF CONTENTS")
    toc_items = [
        ("1", "Executive Summary"),
        ("2", "Key Performance Indicators"),
        ("3", "Financial Analysis"),
        ("4", "Key Findings"),
        ("5", "Strategic Recommendations"),
        ("6", "Action Plan"),
    ]
    for sec in params.get("sections", []):
        toc_items.append(("—", sec.get("title","")))
    for num, item in toc_items:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after  = Pt(1)
        r = p.add_run(f"  {num}.  {item}")
        r.font.size = Pt(10); r.font.color.rgb = DARK_GR
        r.font.name = "Calibri"
    doc.add_page_break()

    # ── Executive Summary ─────────────────────────────────────────────────────
    _h1(doc, "EXECUTIVE SUMMARY")
    exec_sum = params.get("executive_summary", "")
    if exec_sum:
        _parse_content(doc, exec_sum)
    _divider(doc)

    # ── KPI Dashboard ──────────────────────────────────────────────────────────
    fd   = params.get("financial_data", {})
    kpis = fd.get("kpis", [])
    if kpis:
        _h2(doc, "Key Performance Indicators")
        _add_kpi_table(doc, kpis, sym)
        _divider(doc)

    # ── Financial Snapshot ────────────────────────────────────────────────────
    rev = fd.get("revenue", [])
    if rev:
        _h1(doc, "FINANCIAL SNAPSHOT")
        n_periods = min(len(rev), 6)
        labels    = fd.get("period_labels", [f"P{i+1}" for i in range(len(rev))])
        gp        = fd.get("gross_profit",  [0]*len(rev))
        gm_pct    = fd.get("gp_margin",     [0]*len(rev))
        ebitda    = fd.get("ebitda",        [0]*len(rev))
        eb_pct    = fd.get("ebitda_margin", [0]*len(rev))
        net       = fd.get("net_profit",    [0]*len(rev))
        nm_pct    = fd.get("net_margin",    [0]*len(rev))

        hdrs = ["Metric"] + labels[:n_periods]
        rows = [
            ["Revenue"]           + [_fmt(_num(v), sym) for v in rev[:n_periods]],
            ["Gross Profit"]      + [_fmt(_num(v), sym) for v in gp[:n_periods]],
            ["  GP Margin %"]     + [f"{_num(v):.1f}%" for v in gm_pct[:n_periods]],
            ["EBITDA"]            + [_fmt(_num(v), sym) for v in ebitda[:n_periods]],
            ["  EBITDA Margin %"] + [f"{_num(v):.1f}%" for v in eb_pct[:n_periods]],
            ["Net Profit/(Loss)"] + [_fmt(_num(v), sym) for v in net[:n_periods]],
            ["  Net Margin %"]    + [f"{_num(v):.1f}%" for v in nm_pct[:n_periods]],
        ]
        fw = 4.5; cw = (17.0 - fw) / max(1, n_periods)
        _add_table(doc, hdrs, rows, col_widths=[fw]+[cw]*n_periods,
                   highlight_last=True)
        _divider(doc)

    # ── Key Findings ──────────────────────────────────────────────────────────
    findings = params.get("key_findings", [])
    if findings:
        _h1(doc, "KEY FINDINGS")
        for i, f in enumerate(findings, 1):
            if isinstance(f, dict):
                title_f = f.get("title", f"Finding {i}")
                body_f  = f.get("body", "")
            else:
                title_f = f"Finding {i}"; body_f = str(f)
            _finding_box(doc, i, title_f, body_f)
        _divider(doc)

    # ── Main Sections ──────────────────────────────────────────────────────────
    sections = params.get("sections", [])
    for sec in sections:
        level   = sec.get("level", 1)
        s_title = sec.get("title", "")
        content = sec.get("content", "")

        if level == 1:   _h1(doc, s_title)
        elif level == 2: _h2(doc, s_title)
        else:            _h3(doc, s_title)

        if content:
            _parse_content(doc, content)

        # Section KPIs
        sec_kpis = sec.get("kpis", [])
        if sec_kpis:
            _add_kpi_table(doc, sec_kpis, sym)

        # Section tables
        for tbl_data in sec.get("tables", []):
            hdrs = tbl_data.get("headers", [])
            rows = tbl_data.get("rows",    [])
            if hdrs and rows:
                _add_table(doc, hdrs, rows, title=tbl_data.get("title",""))

        _divider(doc)

    # ── Recommendations ───────────────────────────────────────────────────────
    recs = params.get("recommendations", [])
    if recs:
        doc.add_page_break()
        _h1(doc, "STRATEGIC RECOMMENDATIONS")
        for i, rec in enumerate(recs, 1):
            if isinstance(rec, dict):
                t = rec.get("title", f"Recommendation {i}")
                b = rec.get("body", "")
            else:
                t = f"Recommendation {i}"; b = str(rec)
            _rec_box(doc, i, t, b)
        _divider(doc)

    # ── Action Plan ───────────────────────────────────────────────────────────
    action_plan = params.get("action_plan", [])
    if action_plan:
        _h2(doc, "Action Plan")
        hdrs = ["#", "Action", "Owner", "Timeline", "Priority"]
        rows = [[str(i+1)] + [a.get(k,"") for k in
                ["action","owner","timeline","priority"]]
               for i, a in enumerate(action_plan)]
        fw2  = [1.0, 6.5, 3.0, 3.0, 3.5]
        _add_table(doc, hdrs, rows, col_widths=fw2)
        _divider(doc)

    # ── Appendices ────────────────────────────────────────────────────────────
    appendices = params.get("appendices", [])
    if appendices:
        doc.add_page_break()
        _h1(doc, "APPENDICES")
        for app in appendices:
            _h2(doc, app.get("title", "Appendix"))
            if app.get("content"):
                _parse_content(doc, app["content"])
            for tbl_d in app.get("tables", []):
                hdrs = tbl_d.get("headers", [])
                rows = tbl_d.get("rows",    [])
                if hdrs and rows:
                    _add_table(doc, hdrs, rows, title=tbl_d.get("title",""))
            _divider(doc)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()
