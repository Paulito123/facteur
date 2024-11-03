from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.text.run import Font, Run
from docx.dml.color import ColorFormat
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_LEADER, WD_BREAK
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.style import WD_STYLE_TYPE
from helper import convert_to_pdf, hide_table_borders, set_borders_detail_tabel


document = Document()

titel_1_stijl = document.styles['Heading 1']
titel_1_stijl.font.size = Pt(20)
titel_1_stijl.font.name = 'DejaVu Sans Mono'
titel_1_stijl.font.bold = False
titel_1_stijl.font.color.rgb = RGBColor(0, 0, 0)
# titel_1_stijl.paragraph_format.space_after = Pt(0)
titel_1_stijl.paragraph_format.line_spacing = 0

titel_2_stijl = document.styles['Heading 2']
titel_2_stijl.font.size = Pt(14)
titel_2_stijl.font.bold = False
titel_2_stijl.font.color.rgb = RGBColor(0, 0, 0)
tabel_hoofding_stijl = document.styles['Heading 3']
tabel_hoofding_stijl.font.size = Pt(12)
tabel_hoofding_stijl.font.bold = True
tekst_stijl = document.styles['Normal']
tekst_stijl.font.size = Pt(10)
tekst_stijl.font.name = 'DejaVu Sans Mono'
footer_stijl = document.styles['Footer']
footer_stijl.font.name = 'DejaVu Sans Light'
footer_stijl.font.size = Pt(10)

sections = document.sections
for section in sections:
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)

# Create a new header for the first section
sections[0].different_first_page_header_footer = True
header = sections[0].first_page_header

run = header.paragraphs[0].add_run()
run.add_picture('files/images/tokaio.png', width=Cm(7.5))
header.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run.add_break()

hoofd = header.add_table(rows=2, cols=3, width=Cm(19))

rij1 = hoofd.rows[0].cells
titel_cell = rij1[0]
titel_cell.paragraphs[0].style = titel_1_stijl
run = titel_cell.paragraphs[0].add_run()
run.add_text('Factuur')

adres_titel_cell = rij1[2]
adres_titel_cell.paragraphs[0].style = titel_2_stijl
adres_titel_cell.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM
run = adres_titel_cell.paragraphs[0].add_run()
run.add_text('Factuur adres')

rij2 = hoofd.rows[1].cells
details_cell = rij2[0]
run = details_cell.paragraphs[0].add_run()

run.add_text(f'Factuur nr.')
run.add_break(WD_BREAK.LINE)
run.add_text('Factuur datum')
run.add_break(WD_BREAK.LINE)
run.add_text(f'Factuur nr.')
run.add_break(WD_BREAK.LINE)
run.add_text(f'Betalings datum')

run = rij2[2].paragraphs[0].add_run()

run.add_text('Kerkstraat 1')
run.add_break(WD_BREAK.LINE)
run.add_text(f'1234 AB Amsterdam')
run.add_break(WD_BREAK.LINE)
run.add_text(f'Nederland')

hide_table_borders(hoofd)

# Factuur detail
aantal_artikels = 25

detail_tabel_titel = document.add_paragraph('Factuur details')
detail_tabel_titel.style = titel_2_stijl
detail_tabel = document.add_table(rows=aantal_artikels + 2, cols=5)

detail_tabel.autofit = False 
detail_tabel.allow_autofit = False

# HEADER
# 19 cm te verdelen over 5 kolommen (3.8)
# Product beschrijving
detail_tabel.columns[0].width = Cm(6.5)
# detail_tabel.rows[0].cells[0].width = Cm(4.5)
detail_tabel.rows[0].cells[0].text = 'Product beschrijving'

# Aantal
detail_tabel.columns[1].width = Cm(2.5)
detail_tabel.rows[0].cells[1].text = 'Aantal'

# Prijs
detail_tabel.columns[2].width = Cm(3)
detail_tabel.rows[0].cells[2].text = 'Prijs'

# BTW
detail_tabel.columns[3].width = Cm(3.5)
detail_tabel.rows[0].cells[3].text = 'BTW (21%)'

# Bedrag
detail_tabel.columns[4].width = Cm(3.5)
detail_tabel.rows[0].cells[4].text = 'Bedrag'

# DETAILS
for product_nr in range(1, aantal_artikels+1):
    detail_tabel.rows[product_nr].cells[0].text = f'Product {product_nr}'
    detail_tabel.rows[product_nr].cells[1].text = f'{product_nr}'
    detail_tabel.rows[product_nr].cells[2].text = f'€ {100}'
    detail_tabel.rows[product_nr].cells[3].text = f'€ {21}'
    detail_tabel.rows[product_nr].cells[4].text = f'€ {121}'

# Totaal
detail_tabel.rows[aantal_artikels+1].cells[3].paragraphs[0].add_run('Subtotaal')
detail_tabel.rows[aantal_artikels+1].cells[3].paragraphs[0].add_run().add_break(WD_BREAK.LINE)
detail_tabel.rows[aantal_artikels+1].cells[3].paragraphs[0].add_run('BTW')
detail_tabel.rows[aantal_artikels+1].cells[3].paragraphs[0].add_run().add_break(WD_BREAK.LINE)
detail_tabel.rows[aantal_artikels+1].cells[3].paragraphs[0].add_run('Totaal').bold = True

detail_tabel.rows[aantal_artikels+1].cells[4].paragraphs[0].add_run(f'€ {aantal_artikels * 100}').bold = False
detail_tabel.rows[aantal_artikels+1].cells[4].paragraphs[0].add_run().add_break(WD_BREAK.LINE)
detail_tabel.rows[aantal_artikels+1].cells[4].paragraphs[0].add_run(f'€ {aantal_artikels * 21}')
detail_tabel.rows[aantal_artikels+1].cells[4].paragraphs[0].add_run().add_break(WD_BREAK.LINE)
detail_tabel.rows[aantal_artikels+1].cells[4].paragraphs[0].add_run(f'€ {aantal_artikels * 121}').bold = True

set_borders_detail_tabel(detail_tabel)

trailing_text = document.add_paragraph('Hier komt nog wat tekst onder de tabel')

footer = sections[0].first_page_footer
run = footer.paragraphs[0].add_run()
run.add_break()
run.add_text('Tokaio BV')
run.add_break()
run.add_text('Kerkstraat 1')
run.add_break()
run.add_text('Nog wa info ')
run.add_break()
run.add_text('En hoplaaa')

other_footer = sections[0].footer
run = other_footer.paragraphs[0].add_run()
run.add_break()
run.add_text('Tokaio BV')
run.add_break()
run.add_text('Kerkstraat 1')
run.add_break()
run.add_text('Nog wa info ')
run.add_break()
run.add_text('En hoplaaa')

document.add_page_break()

styles = document.styles


document.save('files/invoices/voorbeeld2.docx')
convert_to_pdf('files/invoices/voorbeeld2.docx', 'files/invoices/voorbeeld2.pdf')
