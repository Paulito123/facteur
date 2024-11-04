from docx import Document
from doc_helper import DocHelper


# create a new document
dhelpr = DocHelper()
document = Document()

# set the styles for the document
dhelpr.set_styles(document)

# set the header, body and footer
dhelpr.set_header(document)
dhelpr.set_body(document)
dhelpr.set_footer(document)

# save the document and convert it to pdf
document.save('files/invoices/voorbeeld2.docx')
dhelpr.convert_to_pdf('files/invoices/voorbeeld2.docx', 'files/invoices/voorbeeld2.pdf')
