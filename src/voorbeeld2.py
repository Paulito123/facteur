import json

from docx import Document
from doc_helper import DocHelper



if __name__ == "__main__":
    # create a new document
    dhelpr = DocHelper()
    document = Document()

    # load example data
    data = {}
    with open("files/config/invoice_example.json", "r") as f:
        data = json.load(f)

    # set the styles for the document
    dhelpr.set_styles(document)

    # set the header, body and footer
    dhelpr.set_header(document, data['header'])
    dhelpr.set_body(document, data['body'])
    dhelpr.set_footer(document, data['footer'])

    # save the document and convert it to pdf
    document.save('files/invoices/voorbeeld2.docx')
    dhelpr.convert_to_pdf('files/invoices/voorbeeld2.docx', 'files/invoices/voorbeeld2.pdf')
