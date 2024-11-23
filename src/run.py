import json

from docx import Document
from doc_helper import DocHelper
from doc_processor import DocProcessor
from enumerations import DocumentType, InvoiceTemplate


if __name__ == "__main__":
    data = {}
    
    with open("files/config/prd_arg_202411.json", "r") as f:
        data = json.load(f)

    doc_generator = DocProcessor(data)
    doc_generator.smart_generate(
        DocumentType.INVOICE, 
        invoice_type=InvoiceTemplate.ARGENTA, 
        is_test_run=True
    )

    # with open("files/config/prd_isc.json", "r") as f:
    #     data = json.load(f)
    
    # doc_generator = DocProcessor(data)
    # doc_generator.smart_generate(
    #     DocumentType.INVOICE,
    #     is_test_run=True
    # )
