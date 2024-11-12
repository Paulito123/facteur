import json

from docx import Document
from doc_helper import DocHelper
from doc_processor import DocProcessor
from enumerations import DocumentType



if __name__ == "__main__":
    data = {}
    with open("files/config/smart_example.json", "r") as f:
        data = json.load(f)
    
    doc_generator = DocProcessor(data)
    doc_generator.smart_generate(DocumentType.INVOICE)
