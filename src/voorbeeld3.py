import json

from docx import Document
from doc_helper import DocHelper
from doc_processor import DocProcessor



if __name__ == "__main__":
    data = {}
    with open("files/config/smart.json", "r") as f:
        data = json.load(f)
    
    doc_generator = DocProcessor(data)
    doc_generator.smart_generate()