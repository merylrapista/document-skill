from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn


class DocxSkill:
    name = "docx"
    description = "Generates and manipulates DOCX documents"

    def __init__(self, templates_dir: str | None = None):
        self._doc: Document | None = None
        self._templates_dir = Path(templates_dir) if templates_dir else None

    def execute(self, action: str, **params):
        method = getattr(self, f"_{action}", None)
        if not method:
            raise ValueError(f"Unknown action: {action}")
        return method(**params)

    def _create(self, template: str | None = None) -> dict:
        if template and self._templates_dir:
            template_path = self._templates_dir / f"{template}.docx"
            if template_path.exists():
                self._doc = Document(str(template_path))
                return {"status": "ok", "message": f"Created from template: {template}"}
        self._doc = Document()
        return {"status": "ok", "message": "Created blank document"}

    def _add_heading(self, text: str, level: int = 1) -> dict:
        self._doc.add_heading(text, level=level)
        return {"status": "ok", "element": "heading", "text": text}

    def _add_paragraph(self, text: str, style: str | None = None,
                      bold: bool = False, italic: bool = False,
                      alignment: str | None = None,
                      font_size: int | None = None,
                      font_color: str | None = None) -> dict:
        para = self._doc.add_paragraph()
        run = para.add_run(text)
        run.bold = bold
        run.italic = italic
        if font_size:
            run.font.size = Pt(font_size)
        if font_color:
            run.font.color.rgb = RGBColor(*self._parse_color(font_color))
        if style:
            para.style = self._doc.styles[style]
        if alignment:
            para.alignment = self._parse_alignment(alignment)
        return {"status": "ok", "element": "paragraph", "text": text}

    def _add_table(self, headers: list[str], rows: list[list[str]],
                   style: str = "Table Grid") -> dict:
        table = self._doc.add_table(rows=1, cols=len(headers))
        table.style = self._doc.styles[style]
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, header in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.text = header
            run = cell.paragraphs[0].runs[0] if cell.paragraphs[0].runs else None
            if run:
                run.bold = True
        for row_data in rows:
            row = table.add_row()
            for i, cell_text in enumerate(row_data):
                row.cells[i].text = cell_text
        return {"status": "ok", "element": "table", "rows": len(rows) + 1}

    def _add_list(self, items: list[str], ordered: bool = False) -> dict:
        for item in items:
            if ordered:
                self._doc.add_paragraph(item, style="List Number")
            else:
                self._doc.add_paragraph(item, style="List Bullet")
        return {"status": "ok", "element": "list", "items": len(items)}

    def _add_image(self, image_path: str, width: float | None = None,
                   height: float | None = None) -> dict:
        path = Path(image_path)
        if not path.exists():
            raise FileNotFoundError(f"Image not found: {image_path}")
        kwargs = {}
        if width:
            kwargs["width"] = Inches(width)
        if height:
            kwargs["height"] = Inches(height)
        self._doc.add_picture(str(path), **kwargs)
        return {"status": "ok", "element": "image", "path": image_path}

    def _add_page_break(self) -> dict:
        self._doc.add_page_break()
        return {"status": "ok", "element": "page_break"}

    def _set_header(self, text: str, alignment: str = "center") -> dict:
        section = self._doc.sections[0]
        header = section.header
        para = header.paragraphs[0]
        para.text = text
        para.alignment = self._parse_alignment(alignment)
        return {"status": "ok", "element": "header", "text": text}

    def _set_footer(self, text: str, alignment: str = "center") -> dict:
        section = self._doc.sections[0]
        footer = section.footer
        para = footer.paragraphs[0]
        para.text = text
        para.alignment = self._parse_alignment(alignment)
        return {"status": "ok", "element": "footer", "text": text}

    def _add_page_numbers(self) -> dict:
        section = self._doc.sections[0]
        footer = section.footer
        para = footer.paragraphs[0]
        run = para.add_run()
        fld_char1 = run._r.makeelement(qn("w:fldChar"), {qn("w:fldCharType"): "begin"})
        run._r.append(fld_char1)
        instr_text = run._r.makeelement(qn("w:instrText"), {})
        instr_text.text = " PAGE "
        run._r.append(instr_text)
        fld_char2 = run._r.makeelement(qn("w:fldChar"), {qn("w:fldCharType"): "end"})
        run._r.append(fld_char2)
        return {"status": "ok", "element": "page_numbers"}

    def _set_properties(self, title: str | None = None,
                        author: str | None = None,
                        category: str | None = None) -> dict:
        props = self._doc.core_properties
        if title:
            props.title = title
        if author:
            props.author = author
        if category:
            props.category = category
        return {"status": "ok", "message": "Properties updated"}

    def _set_margins(self, top: float = 1.0, bottom: float = 1.0,
                     left: float = 1.0, right: float = 1.0) -> dict:
        section = self._doc.sections[0]
        section.top_margin = Inches(top)
        section.bottom_margin = Inches(bottom)
        section.left_margin = Inches(left)
        section.right_margin = Inches(right)
        return {"status": "ok", "message": "Margins updated"}

    def _save(self, output_path: str) -> dict:
        path = Path(output_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        self._doc.save(str(path))
        return {"status": "ok", "message": f"Saved to {output_path}", "path": output_path}

    def _parse_color(self, color: str) -> tuple:
        if color.startswith("#"):
            color = color[1:]
        r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
        return (r, g, b)

    def _parse_alignment(self, alignment: str):
        mapping = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        return mapping.get(alignment.lower(), WD_ALIGN_PARAGRAPH.LEFT)
