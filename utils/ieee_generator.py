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
import matplotlib.pyplot as plt
from PIL import Image
import tempfile
from docx.shared import Pt, Inches, RGBColor

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

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
            heading = doc.add_paragraph(f"{idx}. {section_data['heading'].upper()}")
            format_heading(heading)

            if 'subsections' in section_data:
                for sub_idx, sub in enumerate(section_data['subsections'], 1):
                    subheading = doc.add_paragraph(f"{idx}.{sub_idx} {sub['heading']}")
                    format_heading(subheading)

                    if 'content' in sub:
                        content_para = doc.add_paragraph(sub['content'])
                        content_para.paragraph_format.first_line_indent = Inches(0.5)
                        content_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

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
                para = doc.add_paragraph(section_data['content'])
                para.paragraph_format.first_line_indent = Inches(0.5)
                para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # References
        ref_title = doc.add_paragraph("References")
        ref_title.style = "Heading 2"
        for i, ref in enumerate(data['references'], 1):
            doc.add_paragraph(f"[{i}] {ref}")

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
