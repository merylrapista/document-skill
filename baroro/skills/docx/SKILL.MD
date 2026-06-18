# DOCX Generation Skill – Professional Document Builder

Use this skill to generate industry-standard `.docx` documents with `python-docx`.  
Every pattern here follows professional publishing conventions used in consulting, legal, and corporate environments.

## Dependency

```bash
pip install python-docx
```

## Core imports

```python
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement
from docx.enum.dml import WD_THEME_COLOR
import datetime
```

---

## 1. Complete Professional Template

This is the recommended starting point — a full document shell with cover page, section breaks, headers/footers, and page numbering.

```python
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION_START
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement
import datetime


def set_default_font(doc, name='Calibri', size=Pt(11), color=RGBColor(0x33, 0x33, 0x33)):
    style = doc.styles['Normal']
    font = style.font
    font.name = name
    font.size = size
    font.color.rgb = color
    style.paragraph_format.space_after = Pt(6)
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.line_spacing = 1.15


def configure_heading_styles(doc):
    for level, size, color_hex in [
        (1, Pt(20), '1A3C5E'),  # Heading 1 – dark navy
        (2, Pt(15), '2B579A'),  # Heading 2 – medium blue
        (3, Pt(13), '3B6CB4'),  # Heading 3 – lighter blue
    ]:
        style = doc.styles[f'Heading {level}']
        style.font.name = 'Calibri'
        style.font.size = size
        style.font.bold = True
        style.font.color.rgb = RGBColor(*bytes.fromhex(color_hex))
        style.font.space_before = Pt(18 if level == 1 else 12)
        style.font.space_after = Pt(8)


def add_section_break(doc, orientation=WD_SECTION_START.NEW_PAGE):
    """Add a section break (new page) with header/footer decoupled."""
    section = doc.add_section(orientation)
    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False
    return section


def add_header_rule(section):
    """Add a thin horizontal line below the header."""
    p = section.header.paragraphs[0]
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '4')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '999999')
    pBdr.append(bottom)
    for tag in ['w:shd', 'w:tabs', 'w:spacing', 'w:ind', 'w:jc', 'w:rPr']:
        el = pPr.find(qn(tag))
        if el is not None:
            el.addprevious(pBdr)
            break
    else:
        pPr.append(pBdr)


def set_header_content(section, left_text=None, right_text=None):
    """Two-column header: left = document info, right = page number."""
    add_header_rule(section)
    p = section.header.paragraphs[0]
    p.clear()

    if left_text:
        run = p.add_run(left_text)
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        run.font.name = 'Calibri'

    if right_text:
        # Tab to right-aligned position
        tab_stops = p.paragraph_format.tab_stops
        tab_stops.add_tab_stop(Cm(15.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.SPACES)
        run2 = p.add_run(f'\t{right_text}')
        run2.font.size = Pt(8)
        run2.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        run2.font.name = 'Calibri'


def set_footer_content(section, left_text=None, center_text=None):
    """Professional footer with document info and page numbers."""
    p = section.footer.paragraphs[0]
    p.clear()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if left_text:
        run = p.add_run(left_text)
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
        run.font.name = 'Calibri'
        p.add_run('\t')


def add_page_number_field(paragraph):
    """Insert dynamic page-number field (Page X of Y)."""
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar1)

    run2 = paragraph.add_run()
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = ' PAGE '
    run2._r.append(instrText)

    run3 = paragraph.add_run()
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run3._r.append(fldChar2)

    run4 = paragraph.add_run(' of ')

    run5 = paragraph.add_run()
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'begin')
    run5._r.append(fldChar3)

    run6 = paragraph.add_run()
    instrText2 = OxmlElement('w:instrText')
    instrText2.set(qn('xml:space'), 'preserve')
    instrText2.text = ' NUMPAGES '
    run6._r.append(instrText2)

    run7 = paragraph.add_run()
    fldChar4 = OxmlElement('w:fldChar')
    fldChar4.set(qn('w:fldCharType'), 'end')
    run7._r.append(fldChar4)


def set_page_number(section, style='DECIMAL', start=1):
    """Set page numbering for a section (roman/arabic)."""
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn('w:pgNumType'))
    if pgNumType is None:
        pgNumType = OxmlElement('w:pgNumType')
        cols = sectPr.find(qn('w:cols'))
        docGrid = sectPr.find(qn('w:docGrid'))
        if cols is not None:
            cols.addprevious(pgNumType)
        elif docGrid is not None:
            docGrid.addprevious(pgNumType)
        else:
            sectPr.append(pgNumType)
    pgNumType.set(qn('w:fmt'), style)  # 'decimal', 'upperRoman', 'lowerRoman'
    pgNumType.set(qn('w:start'), str(start))


def add_cover_page(doc, title, subtitle=None, author=None, date=None, color_hex='1A3C5E'):
    """Professional cover page with centered title block."""
    # Blank paragraphs for vertical spacing
    for _ in range(6):
        doc.add_paragraph()

    # Horizontal rule line
    p_line = doc.add_paragraph()
    p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = p_line._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '18')
    bottom.set(qn('w:space'), '6')
    bottom.set(qn('w:color'), color_hex)
    pBdr.append(bottom)
    for tag in ['w:shd', 'w:tabs', 'w:spacing', 'w:ind', 'w:jc', 'w:rPr']:
        el = pPr.find(qn(tag))
        if el is not None:
            el.addprevious(pBdr)
            break
    else:
        pPr.append(pBdr)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(12)
    run = p_title.add_run(title)
    run.font.size = Pt(32)
    run.font.bold = True
    run.font.color.rgb = RGBColor(*bytes.fromhex(color_hex))
    run.font.name = 'Calibri'

    # Subtitle
    if subtitle:
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p_sub.add_run(subtitle)
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        run.font.name = 'Calibri'

    # Bottom rule
    p_line2 = doc.add_paragraph()
    p_line2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr2 = p_line2._p.get_or_add_pPr()
    pBdr2 = OxmlElement('w:pBdr')
    top = OxmlElement('w:top')
    top.set(qn('w:val'), 'single')
    top.set(qn('w:sz'), '18')
    top.set(qn('w:space'), '6')
    top.set(qn('w:color'), color_hex)
    pBdr2.append(top)
    for tag in ['w:shd', 'w:tabs', 'w:spacing', 'w:ind', 'w:jc', 'w:rPr']:
        el = pPr2.find(qn(tag))
        if el is not None:
            el.addprevious(pBdr2)
            break
    else:
        pPr2.append(pBdr2)

    # Spacer
    for _ in range(4):
        doc.add_paragraph()

    # Metadata block
    meta_items = []
    if author:
        meta_items.append(('Prepared by', author))
    if date:
        meta_items.append(('Date', date))
    else:
        meta_items.append(('Date', datetime.date.today().strftime('%B %d, %Y')))

    for label, value in meta_items:
        p_meta = doc.add_paragraph()
        p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_l = p_meta.add_run(f'{label}:  ')
        run_l.font.size = Pt(10)
        run_l.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        run_l.font.name = 'Calibri'
        run_v = p_meta.add_run(value)
        run_v.font.size = Pt(11)
        run_v.font.bold = True
        run_v.font.color.rgb = RGBColor(*bytes.fromhex(color_hex))
        run_v.font.name = 'Calibri'


def add_toc(doc, title='Table of Contents'):
    """Insert a Word TOC field (updatable in Word)."""
    doc.add_paragraph()
    h = doc.add_heading(title, level=1)

    # TOC field
    p = doc.add_paragraph()
    run = p.add_run()
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar)

    run2 = p.add_run()
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = ' TOC \\o "1-3" \\h \\z \\u '
    run2._r.append(instrText)

    run3 = p.add_run()
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    run3._r.append(fldChar2)

    # Placeholder text
    run4 = p.add_run('[Right-click → Update Field to generate Table of Contents]')
    run4.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
    run4.font.size = Pt(10)
    run4.font.italic = True

    run5 = p.add_run()
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    run5._r.append(fldChar3)

    doc.add_paragraph()


def add_watermark_text(doc, text='DRAFT'):
    """Add a diagonal watermark text to every page."""
    for section in doc.sections:
        header = section.header
        p = header.paragraphs[0]
        run = p.add_run()
        rPr = run._r.get_or_add_rPr()
        # Make text transparent and large
        rPr.append(parse_xml(
            f'<w:rPr {nsdecls("w")}>'
            f'  <w:sz w:val="720"/>'
            f'  <w:color w:val="CCCCCC"/>'
            f'  <w:rFonts w:ascii="Calibri"/>'
            f'</w:rPr>'
        ))
        # Add drawing for rotated text
        drawing = parse_xml(
            f'<w:drawing {nsdecls("w", "wp", "a", "r")}>'
            f'  <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264"'
            f'    behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">'
            f'    <wp:simplePos x="0" y="0"/>'
            f'    <wp:positionH relativeFrom="page">'
            f'      <wp:posOffset>0</wp:posOffset>'
            f'    </wp:positionH>'
            f'    <wp:positionV relativeFrom="page">'
            f'      <wp:posOffset>0</wp:posOffset>'
            f'    </wp:positionV>'
            f'    <wp:extent cx="914400" cy="914400"/>'
            f'    <wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'    <wp:wrapNone/>'
            f'    <wp:docPr id="1" name="Watermark"/>'
            f'    <a:graphic>'
            f'      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
            f'        <wsp:shape xmlns:wsp="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">'
            f'          <wsp:spPr>'
            f'            <a:xfrm>'
            f'              <a:off x="0" y="0"/>'
            f'              <a:ext cx="914400" cy="914400"/>'
            f'            </a:xfrm>'
            f'            <a:prstGeom prst="rect"/>'
            f'            <a:noFill/>'
            f'          </wsp:spPr>'
            f'          <wsp:txBody>'
            f'            <a:bodyPr rot="-5400000"/>'
            f'            <a:lstStyle/>'
            f'            <a:p>'
            f'              <a:pPr algn="ctr"/>'
            f'              <a:r>'
            f'                <a:rPr sz="720" b="1"/>'
            f'                <a:t>{text}</a:t>'
            f'              </a:r>'
            f'            </a:p>'
            f'          </wsp:txBody>'
            f'        </wsp:shape>'
            f'      </a:graphicData>'
            f'    </a:graphic>'
            f'  </wp:anchor>'
            f'</w:drawing>'
        )
        run._r.append(drawing)
```

---

## 2. Usage Blueprint — Putting It All Together

```python
# ── Build document ──
doc = Document()

# ── Configure defaults ──
set_default_font(doc)
configure_heading_styles(doc)

# ── Section 0: Cover page (no header/footer) ──
section0 = doc.sections[0]
section0.header.is_linked_to_previous = False
section0.footer.is_linked_to_previous = False
# Remove header/footer content for cover
section0.header.paragraphs[0].clear()
section0.footer.paragraphs[0].clear()
set_page_number(section0, 'upperRoman', 1)

add_cover_page(
    doc,
    title='Strategic Roadmap 2026',
    subtitle='Digital Transformation Initiative',
    author='Strategy & Operations',
    date='June 18, 2026',
)

# ── Section 1: TOC + Body ──
section1 = add_section_break(doc)
set_page_number(section1, 'decimal', 1)
set_header_content(section1, left_text='Strategic Roadmap 2026', right_text='Page ')
add_page_number_field(section1.header.paragraphs[0])
set_footer_content(section1, left_text='Confidential')

# TOC
add_toc(doc)
doc.add_page_break()

# ── Content sections ──
doc.add_heading('Executive Summary', level=1)
doc.add_paragraph(
    'This document outlines the strategic roadmap for the Digital Transformation Initiative '
    'spanning fiscal years 2026 through 2028. The initiative aims to modernize core infrastructure, '
    'streamline operations, and enhance customer experience through technology-driven solutions.'
)

doc.add_heading('Scope & Objectives', level=1)
doc.add_heading('Primary Objectives', level=2)
doc.add_paragraph('The following objectives have been identified as critical to the success of the initiative:')

objectives = [
    'Migrate legacy systems to cloud-native architecture',
    'Implement AI-driven process automation across operations',
    'Establish a unified data platform for real-time analytics',
    'Deploy customer-facing digital portals with omni-channel support',
]
for obj in objectives:
    p = doc.add_paragraph(obj, style='List Bullet')

doc.add_heading('Key Milestones', level=2)

# Professional table
table = doc.add_table(rows=6, cols=4, style='Light Shading Accent 1')
table.alignment = WD_TABLE_ALIGNMENT.CENTER

headers = ['Milestone', 'Target Date', 'Owner', 'Status']
for i, hdr in enumerate(headers):
    cell = table.rows[0].cells[i]
    cell.text = hdr
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for p in cell.paragraphs:
        for run in p.runs:
            run.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    # Dark header background
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="1A3C5E" w:val="clear"/>')
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ['w:noWrap', 'w:tcMar']:
        el = tcPr.find(qn(tag))
        if el is not None:
            el.addprevious(shading)
            break
    else:
        tcPr.append(shading)

milestones = [
    ('Cloud Infrastructure Setup', 'Q1 2026', 'IT Operations', 'In Progress'),
    ('Data Platform MVP', 'Q2 2026', 'Data Engineering', 'On Track'),
    ('Process Automation Rollout', 'Q3 2026', 'Business Ops', 'Planning'),
    ('Customer Portal v1', 'Q4 2026', 'Product', 'Not Started'),
    ('Full Deployment & Review', 'Q1 2027', 'Program Office', 'Not Started'),
]
for r, (m, d, o, s) in enumerate(milestones, 1):
    table.rows[r].cells[0].text = m
    table.rows[r].cells[1].text = d
    table.rows[r].cells[2].text = o
    table.rows[r].cells[3].text = s
    # Alternate row shading
    if r % 2 == 0:
        for cell in table.rows[r].cells:
            shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F2F5FA" w:val="clear"/>')
            cell._tc.get_or_add_tcPr().append(shading)

# ── Save ──
doc.save('output.docx')
```

---

## 3. Professional Table Patterns

### Header row with background color
```python
for cell in table.rows[0].cells:
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="1A3C5E" w:val="clear"/>')
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ['w:noWrap', 'w:tcMar']:
        el = tcPr.find(qn(tag))
        if el is not None:
            el.addprevious(shading)
            break
    else:
        tcPr.append(shading)
    for p in cell.paragraphs:
        for run in p.runs:
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.bold = True
```

### Alternating row colors
```python
for i, row in enumerate(table.rows[1:], 1):
    if i % 2 == 0:
        for cell in row.cells:
            shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F2F5FA" w:val="clear"/>')
            tcPr = cell._tc.get_or_add_tcPr()
            for tag in ['w:noWrap', 'w:tcMar']:
                el = tcPr.find(qn(tag))
                if el is not None:
                    el.addprevious(shading)
                    break
            else:
                tcPr.append(shading)
```

### Cell vertical alignment + padding
```python
cell = table.cell(0, 0)
cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
tc = cell._tc
tcPr = tc.get_or_add_tcPr()
tcMar = OxmlElement('w:tcMar')
for edge in ('top', 'start', 'bottom', 'end'):
    m = OxmlElement(f'w:{edge}')
    m.set(qn('w:w'), '57')  # ~1mm
    m.set(qn('w:type'), 'dxa')
    tcMar.append(m)
tcPr.append(tcMar)
```

### Merge cells
```python
table.cell(0, 0).merge(table.cell(0, 2))
```

---

## 4. Advanced Styling

### Multi-level numbered list
```python
from docx.oxml import OxmlElement

def add_numbered_paragraph(doc, text, level=0):
    p = doc.add_paragraph(text, style='List Number')
    pPr = p._p.get_or_add_pPr()
    ind = OxmlElement('w:ind')
    ind.set(qn('w:left'), str(720 + level * 720))
    ind.set(qn('w:hanging'), '360')
    pPr.append(ind)
    numPr = OxmlElement('w:numPr')
    ilvl = OxmlElement('w:ilvl')
    ilvl.set(qn('w:val'), str(level))
    numPr.append(ilvl)
    pPr.append(numPr)
    return p
```

### Side headings (book-style)
```python
def add_side_heading(doc, text, color_hex='1A3C5E'):
    """Heading left-aligned in a 1-row table for a clean sidebar look."""
    table = doc.add_table(rows=1, cols=2)
    table.columns[0].width = Cm(3.5)
    table.columns[1].width = Cm(13)
    cell_l = table.cell(0, 0)
    p = cell_l.paragraphs[0]
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(*bytes.fromhex(color_hex))
    run.font.name = 'Calibri'
    # Right-aligned vertical bar
    cell_r = table.cell(0, 1)
    sdt = parse_xml(
        f'<w:sdt {nsdecls("w")}>'
        f'  <w:sdtPr>'
        f'    <w:rPr/>'
        f'  </w:sdtPr>'
        f'  <w:sdtContent>'
        f'    <w:p><w:r><w:rPr><w:color w:val="{color_hex}"/><w:sz w:val="20"/></w:rPr>'
        f'    <w:t>▎</w:t></w:r></w:p>'
        f'  </w:sdtContent>'
        f'</w:sdt>'
    )
    cell_r._tc.append(sdt)
    doc.add_paragraph()  # spacer
```

### Callout / note box
```python
def add_callout(doc, text, bg_color='E8F0FE', border_color='1A3C5E'):
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    cell.text = text
    cell.paragraphs[0].style = doc.styles['Normal']
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    # Background
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_color}" w:val="clear"/>')
    tcPr.append(shading)
    # Border
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'start', 'bottom', 'end'):
        border = OxmlElement(f'w:{edge}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '8')
        border.set(qn('w:color'), border_color)
        tcBorders.append(border)
    tcPr.append(tcBorders)
    return table
```

### Page border (entire page)
```python
def add_page_border(section, color='CCCCCC', size='12'):
    sectPr = section._sectPr
    pgBorders = OxmlElement('w:pgBorders')
    pgBorders.set(qn('w:offsetFrom'), 'page')
    for edge in ('top', 'left', 'bottom', 'right', 'start', 'end'):
        border = OxmlElement(f'w:{edge}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), size)
        border.set(qn('w:space'), '24')
        border.set(qn('w:color'), color)
        pgBorders.append(border)
    sectPr.append(pgBorders)
```

---

## 5. Complete Reference – All Imports

```python
# Shared
from docx.shared import Inches, Pt, Cm, RGBColor, Emu, Twips

# Document
from docx import Document
from docx.section import Section
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement

# Paragraphs
from docx.enum.text import (
    WD_ALIGN_PARAGRAPH,
    WD_LINE_SPACING,
    WD_TAB_ALIGNMENT,
    WD_TAB_LEADER,
    WD_UNDERLINE,
    WD_BREAK,
)

# Tables
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL

# Sections
from docx.enum.section import WD_ORIENT, WD_SECTION_START

# Styles
from docx.enum.style import WD_STYLE_TYPE, WD_STYLE

# Drawing / Images
from docx.enum.dml import WD_THEME_COLOR

# Datetime for auto-date
import datetime
```

---

## 6. Design Guidelines

### Color palette (corporate conservative)
| Usage | Hex | Description |
|-------|-----|-------------|
| Primary text | `#333333` | Body text |
| Heading 1 | `#1A3C5E` | Dark navy |
| Heading 2 | `#2B579A` | Medium blue |
| Heading 3 | `#3B6CB4` | Accent blue |
| Table header | `#1A3C5E` white text | Header row |
| Table alt row | `#F2F5FA` | Alternating |
| Callout bg | `#E8F0FE` | Highlight box |
| Subtle lines | `#CCCCCC` | Borders/rules |
| Muted text | `#999999` | Footer/metadata |

### Font stack
- **Body:** Calibri 11pt, 1.15 line spacing, dark gray (#333)
- **Headings:** Calibri Bold in navy/blue hierarchy
- **Code/monospace:** Consolas 9pt (if needed)

### Page layout (A4)
- **Page:** A4 (21.0 × 29.7 cm)
- **Margins:** 2.54 cm all sides (standard US letter) or 2.0 cm (sleeker)
- **Header:** 1.25 cm from edge
- **Footer:** 1.0 cm from edge

### Spacing rhythm
- After heading: 8 pt
- After paragraph: 6 pt
- Before heading 1: 18 pt
- Before heading 2: 12 pt
- Before heading 3: 10 pt

### When to use which pattern
- **Short documents (1–5 pp):** Cover page unnecessary — start directly with heading + content, standard header/footer
- **Reports (6–20 pp):** Cover page + TOC + section headers/footers
- **Proposals / contracts:** Add watermark, page borders, TOC, cover page, different headers per section
- **Multi-section:** Each major section gets its own section break so headers can reflect different topics
- **Tables:** Always style header row with dark bg + white bold text, alternating row fills for readability
- **Lists:** Use Word styles (`'List Bullet'`, `'List Number'`) — do not manually add bullet characters

### DOs
- Always call `set_default_font()` and `configure_heading_styles()` early
- Decouple header/footer from previous section on every section break
- Use dynamic fields (PAGE / NUMPAGES / TOC) instead of hard-coded numbers
- Add alt-row shading on any table with more than 3 data rows
- Use `doc.save()` at the very end

### DON'Ts
- Don't mix fonts — stick to Calibri or the corporate font throughout
- Don't hard-code page numbers — use Word field codes
- Don't add empty paragraphs for spacing — use `paragraph_format.space_before/after`
- Don't use `doc.add_heading(level=0)` for cover titles — use custom formatting
- Don't put content on the cover page section — keep it isolated
