import datetime

from typing import Dict, AnyStr
from docx import Document
from docx.shared import Cm
from enumerations import DocumentType, InvoiceTemplate, OfferTemplate, BorderTemplate
from json import load

from config import Config
from doc_helper import DocHelper


class DocProcessor:

    def __init__(self):
        self.db = None

        try:
            with open(Config.PATH_DB, "r") as f:
                self.db = load(f)
        except Exception as e:
            Exception(f"Could not initialize DB: {e}")
        

    def __check_data(self, data) -> bool:
        """
        Validates the provided data dictionary to ensure all required fields are present.
        Args:
            data (dict): The data dictionary to be validated. It should contain the following keys:
                - "delivery_date" (str): The delivery date.
                - "items" (list): A list of items, where each item is a dictionary containing:
                    - "description" (str): Description of the item.
                    - "qty" (int): Quantity of the item.
                    - "vat_pct" (float): VAT percentage for the item.
                    - "base_amt" (float): Base amount for the item.
                - "debtor_id" (str, optional): The ID of the debtor.
                - "debtor_details" (dict, optional): Details of the debtor, containing:
                    - "name" (str): Name of the debtor.
                    - "street" (str): Street address of the debtor.
                    - "number" (str): Street number of the debtor.
                    - "city" (str): City of the debtor.
                    - "zip" (str): ZIP code of the debtor.
                    - "country" (str): Country of the debtor.
                    - "vat" (str): VAT number of the debtor.
                    - "bank_account" (str): Bank account of the debtor.
                    - "rpr" (str): RPR of the debtor.
        Returns:
            bool: True if all required fields are present, False otherwise.
        Raises:
            Exception: If any required field is missing, an exception is raised with a message indicating the missing field.
        """
        
        try:
            # level 1 check
            for s in ["delivery_date", "items"]:
                if s not in data:
                    raise Exception(f"Missing field: {s}")
            
            # conditional checks for debtor
            if not "debtor_id" in data or not "debtor_details" in data:
                raise Exception("Missing field: debtor_id or debtor_details")
            elif not "debtor_id" in data and "debtor_details" in data:
                for s in ["name", "street", "number", "city", "zip", "country", "vat", "bank_account", "rpr"]:
                    if s not in data["debtor_details"]:
                        raise Exception(f"Missing field: {s}")
            
            # conditional checks items
            for s in ["description", "qty", "vat_pct", "base_amt"]:
                if not all(s in item for item in data["items"]):
                    raise Exception(f"Missing field: {s}")
            
            return True
            
        except Exception as e:
            print(e)
            return False
    

    def __generate(self, 
                   data: Dict,
                   doc_type: DocumentType = DocumentType.INVOICE, 
                   invoice_type: InvoiceTemplate = InvoiceTemplate.NEON) -> None:
        """
        fucntion to generate the actual document
        """
        # Check if the data is valid
        if not self.__check_data(data):
            return None
        
        creditor_id = data["creditor_id"]
        debtor_id = data["debtor_id"]
        debtor_person_id = data["debtor_person_id"]
        invoice_date = data["invoice_date"]
        due_date = data["due_date"]
        items = data["items"]

        if doc_type == DocumentType.INVOICE:
            if invoice_type == InvoiceTemplate.NEON:
                self.generate_neon_invoice(data)
            elif invoice_type == InvoiceTemplate.ARGENTA:
                self.generate_argenta_invoice(data)
            else:
                raise Exception("Invalid invoice type")
            
        elif doc_type == DocumentType.OFFER:
            ...
        # Create a new Word document
        document = Document()
        document.add_heading('Invoice', 0)

        return None


    def smart_generate(self, 
                       data: Dict,
                       doc_type: DocumentType = DocumentType.INVOICE, 
                       invoice_type: InvoiceTemplate = InvoiceTemplate.NEON) -> None:
        
        # Check if the data is valid
        if not self.__check_data(data):
            return None
        
        doc_data = {
            "path_image": 'files/images/tokaio.png',
        }
        
        # all the default selection logic goes here
        if doc_type == DocumentType.INVOICE:
            
            doc_data["titel"] = "Factuur"

            if "invoice_date" in data:
                doc_data["invoice_date"] = data["invoice_date"]
            else:
                doc_data["invoice_date"] = datetime.now().strftime("%d-%m-%Y")
            
            if "delivery_date" in data:
                doc_data["delivery_date"] = data["delivery_date"]
            else:
                doc_data["delivery_date"] = datetime.now().strftime("%d-%m-%Y")
            
            policy = self.db["policies"][data["defaults"]["policy_id"]]
            
            currency = self.db["currencies"][data["defaults"]["currency_id"]]
            doc_data["symbol"] = currency["symbol"]

            creditor = self.db["companies"][data["defaults"]["creditor_id"]]
            for key in creditor.keys():
                if key not in ["last_sequences"]:
                    doc_data['creditor_' + key] = creditor[key]
                else:
                    doc_data['invoice_nr'] = f'{datetime.now().year}-{creditor['last_sequences']["invoice"] + 1}'
            

            self.generate_invoice(data)

        elif doc_type == DocumentType.OFFER:
            doc_data["titel"] = "Offerte"
            self.generate_offer(data)


    def generate_offer(self, data: Dict):
        
        # Create a 
        self.__generate(data)

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


    def generate_invoice(self, data: Dict):
        
        # Create a 
        self.__generate(data)


if __name__ == "__main__":
    data = {
        "debtor_id": 3,
        "items": {
            "1": {
                "description": "Neon design", 
                "qty": 1, 
                "vat_pct": 0.21, 
                "base_amt": 25
            },
            "2": {
                "description": "Light board - outdoor - dimmer - remote - 100x34 cm", 
                "qty": 1, 
                "vat_pct": 0.21,
                "base_amt": 429.55
            }
        },
        "delivery_date": "13-11-2024"
    }
    
    doc_generator = DocProcessor()
    doc_generator.smart_generate(data=data)
    