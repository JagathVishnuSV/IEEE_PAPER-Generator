from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.shared import RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO
import base64
import logging
import re
import matplotlib.pyplot as plt
from PIL import Image
import tempfile
from docx.shared import Pt, Inches, RGBColor

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def to_roman(num):
    roman_map = [
        (1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'),
        (100, 'C'), (90, 'XC'), (50, 'L'), (40, 'XL'),
        (10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')
    ]
    result = ""
    for val, sym in roman_map:
        while num >= val:
            result += sym
            num -= val
    return result

def generate_latex_formula_image(latex_code: str) -> str | None:
    try:
        fig, ax = plt.subplots(figsize=(2, 0.5))
        ax.text(0.5, 0.5, f"${latex_code}$", fontsize=14, ha='center', va='center')
        ax.axis('off')
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            plt.savefig(tmp.name, format='png', bbox_inches='tight', transparent=True)
            plt.close(fig)
            return tmp.name
    except Exception as e:
        logger.error(f"Formula rendering failed: {e}")
        return None


logger = logging.getLogger(__name__)
def add_hyperlinks(paragraph, text):
    hyperlink_pattern = r'\[([^\]]+)\]\((https?://[^\)]+)\)'
    last_idx = 0
    for match in re.finditer(hyperlink_pattern, text):
        start, end = match.span()
        if start > last_idx:
            paragraph.add_run(text[last_idx:start])
        run = paragraph.add_run(match.group(1))
        run.font.color.rgb = RGBColor(0, 0, 255)
        run.font.underline = True
        last_idx = end
    if last_idx < len(text):
        paragraph.add_run(text[last_idx:])

def insert_footnotes(paragraph, content):
    footnote_pattern = r"\[\[footnote:(.*?)\]\]"
    matches = re.findall(footnote_pattern, content)
    for note in matches:
        content = content.replace(f"[[footnote:{note}]]", f"[*] {note}")
    paragraph.add_run(content)

def insert_appendix(doc, appendix_data):
    if not appendix_data:
        return
    doc.add_paragraph("Appendix", style="Heading 2")
    for idx, item in enumerate(appendix_data, 1):
        para = doc.add_paragraph(f"{chr(64+idx)}. {item}")
        para.paragraph_format.first_line_indent = Inches(0.5)
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

def generate_ieee_paper(data: dict) -> bytes:
    try:
        doc = Document()

        normal_style = doc.styles['Normal']
        normal_style.font.name = 'Times New Roman'
        normal_style.font.size = Pt(10)

        section = doc.sections[0]
        section.page_width = Inches(8.5)
        section.page_height = Inches(11)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)

        set_single_column_layout(section)

        # Title
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(data['title'].upper())
        run.bold = True
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(0, 0, 0)

        for line in [", ".join(data['authors']),
                     "; ".join(data['affiliations']),
                     ", ".join(data['emails'])]:
            para = doc.add_paragraph(line)
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_paragraph()

        # New section for two-column
        new_section = doc.add_section(start_type=0)
        set_ieee_column_layout(new_section)

        # Abstract
        doc.add_paragraph("Abstract", style="Heading 2")
        abs_para = doc.add_paragraph(data['abstract'])
        abs_para.paragraph_format.first_line_indent = Inches(0.5)
        abs_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # Keywords
        doc.add_paragraph("Keywords", style="Heading 2")
        kw_para = doc.add_paragraph(", ".join(data['keywords']))
        kw_para.paragraph_format.first_line_indent = Inches(0.5)
        kw_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        figure_count = 1
        table_count = 1

        for idx, section_data in enumerate(data['sections'], 1):
            # Section Heading
            roman_idx = to_roman(idx)
            heading = doc.add_paragraph(f"{roman_idx}. {section_data['heading'].upper()}")

            if 'subsections' in section_data:
                for sub_idx, sub in enumerate(section_data['subsections'], 1):
                    alpha = chr(64 + sub_idx)  # A, B, C...
                    subheading = doc.add_paragraph(f"{to_roman(idx)}.{alpha} {sub['heading']}")


                    if 'content' in sub:
                        content_para = doc.add_paragraph()
                        content_para.paragraph_format.first_line_indent = Inches(0.5)
                        content_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        insert_footnotes(content_para, sub['content'])
                        add_hyperlinks(content_para, sub['content'])


                    for img in sub.get('images', []):
                        try:
                            img_stream = BytesIO(base64.b64decode(img['data']))
                            doc.add_picture(img_stream, width=Inches(3))
                            cap = doc.add_paragraph(f"Fig. {figure_count}: {img['caption']}")
                            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            figure_count += 1
                        except Exception as e:
                            logger.error(f"Failed to add image: {e}")

                    for table_data in sub.get('tables', []):
                        try:
                            table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                            table.style = 'Table Grid'
                            for r, row in enumerate(table_data):
                                for c, val in enumerate(row):
                                    table.cell(r, c).text = str(val)
                            doc.add_paragraph(f"Table {table_count}: Data Table").alignment = WD_ALIGN_PARAGRAPH.CENTER
                            table_count += 1
                        except Exception as e:
                            logger.error(f"Failed to render table: {e}")

                    for f_idx, formula in enumerate(sub.get('formulas', []), 1):
                        img_path = generate_latex_formula_image(formula)
                        if img_path:
                            doc.add_picture(img_path, width=Inches(2))
                            caption = doc.add_paragraph(f"Equation {idx}.{sub_idx}.{f_idx}: {formula}")
                            caption.alignment = WD_ALIGN_PARAGRAPH.CENTER

            elif 'content' in section_data:
                para = doc.add_paragraph()
                para.paragraph_format.first_line_indent = Inches(0.5)
                para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                insert_footnotes(para, section_data['content'])
                add_hyperlinks(para, section_data['content'])


        # References
        ref_title = doc.add_paragraph("References")
        ref_title.style = "Heading 2"
        for i, ref in enumerate(data['references'], 1):
            doc.add_paragraph(f"[{i}] {ref}")

        if 'appendix' in data:
            insert_appendix(doc, data['appendix'])

        buffer = BytesIO()
        doc.save(buffer)
        return buffer.getvalue()

    except Exception as e:
        logger.error(f"IEEE document generation failed: {e}")
        raise RuntimeError(f"Failed to generate document: {e}")


def set_single_column_layout(section):
    sectPr = section._sectPr
    for col in sectPr.xpath('./w:cols'):
        sectPr.remove(col)
    cols = OxmlElement('w:cols')
    cols.set(qn('w:num'), '1')
    sectPr.append(cols)


def set_ieee_column_layout(section):
    sectPr = section._sectPr
    for col in sectPr.xpath('./w:cols'):
        sectPr.remove(col)
    cols = OxmlElement('w:cols')
    cols.set(qn('w:num'), '2')
    cols.set(qn('w:space'), '720')
    sectPr.append(cols)


def format_heading(paragraph):
    run = paragraph.runs[0]
    run.bold = True
    run.font.color.rgb = RGBColor(0, 0, 0)
    paragraph.paragraph_format.space_before = Pt(8)
    paragraph.paragraph_format.space_after = Pt(4)
