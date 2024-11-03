from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Pt
import subprocess


def set_cell_border(cell, **kwargs):
    """
    Set cell`s border
    Usage:

    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
        bottom={"sz": 12, "color": "#00FF00", "val": "single"},
        start={"sz": 24, "val": "dashed", "shadow": "true"},
        end={"sz": 12, "val": "dashed"},
    )
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # check for tag existnace, if none found, then create one
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    # list over all available tags
    for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)

            # check for tag existnace, if none found, then create one
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)

            # looks like order of attributes is important
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))

def hide_table_borders(table):
    # Make table borders invisible
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(
                cell,
                top={"sz": 1, "val": "single", "color": "#FFFFFF", "space": "0"},
                bottom={"sz": 1, "val": "single", "color": "#FFFFFF", "space": "0"},
                start={"sz": 1, "val": "single", "color": "#FFFFFF", "space": "0"},
                end={"sz": 1, "val": "single", "color": "#FFFFFF", "space": "0"},
            )


def set_borders_detail_tabel(table):
    aantal_rijen = len(table.rows)
    counter = 1
    # Set borders for table
    for row in table.rows:
        if counter == 1:
            for cell in row.cells:
                set_cell_border(
                    cell,
                    top={"sz": 6, "val": "single", "color": "#000000", "space": "0"},
                    bottom={"sz": 6, "val": "single", "color": "#000000", "space": "0"},
                    start={"sz": 1, "val": "single", "color": "#FFFFFF", "space": "0"},
                    end={"sz": 1, "val": "single", "color": "#FFFFFF", "space": "0"},
                )
        elif counter != aantal_rijen:
            for cell in row.cells:
                set_cell_border(
                    cell,
                    top={"sz": 1, "val": "single", "color": "#FFFFFF", "space": "0"},
                    bottom={"sz": 1, "val": "single", "color": "#FFFFFF", "space": "0"},
                    start={"sz": 1, "val": "single", "color": "#FFFFFF", "space": "0"},
                    end={"sz": 1, "val": "single", "color": "#FFFFFF", "space": "0"},
                )
        else:
            for cell in row.cells:
                set_cell_border(
                    cell,
                    top={"sz": 6, "val": "single", "color": "#000000", "space": "0"},
                    bottom={"sz": 1, "val": "single", "color": "#FFFFFF", "space": "0"},
                    start={"sz": 1, "val": "single", "color": "#FFFFFF", "space": "0"},
                    end={"sz": 1, "val": "single", "color": "#FFFFFF", "space": "0"},
                )
        counter += 1


def set_tokaio_paragraph_style(document: Document):
    style = document.styles.add_style('Indent', WD_STYLE_TYPE.PARAGRAPH)
    paragraph_format = style.paragraph_format
    paragraph_format.left_indent = Cm(0.5)
    paragraph_format.first_line_indent = Cm(-0.5)
    paragraph_format.space_before = Pt(12)
    paragraph_format.widow_control = True


def convert_to_pdf(docx_path, pdf_path):
    subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', '/'.join(pdf_path.split('/')[:-1]), docx_path])

