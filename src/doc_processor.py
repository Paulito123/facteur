from typing import Dict, AnyStr
from docx import Document
from docx.shared import Cm
from enumerations import DocumentType, InvoiceTemplate, OfferTemplate, BorderTemplate


class DocProcessor:

    def __init__(self, data: Dict = None):
        self.data = data
        

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
    

    def __generate(self, data: Dict, config: Dict = {}) -> None:
        
        # Check if the data is valid
        if not self.__check_data(data):
            return None
        
        creditor_id = data["creditor_id"]
        debtor_id = data["debtor_id"]
        debtor_person_id = data["debtor_person_id"]
        invoice_date = data["invoice_date"]
        due_date = data["due_date"]
        items = data["items"]

        # Create a new Word document
        document = Document()
        document.add_heading('Invoice', 0)

        return None


    def smart_generate(self, 
                       doc_type: DocumentType = DocumentType.INVOICE, 
                       invoice_type: InvoiceTemplate = InvoiceTemplate.NEON) -> None:
        ...
        

    def generate_invoice_by_id(self, id: AnyStr):
        pass

    def generate_offer_by_id(self, id: AnyStr):
        pass

    def generate_argenta_invoice(self, data: Dict):
        
        # Generate the invoice
        return self.__generate(data)
        
    def generate_neon_invoice(self, data: Dict):
        
        # Generate the invoice
        return self.__generate(data)

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
        "delivery_date": "2024-11-12"
    }
    
    doc_generator = DocProcessor(data=data)
    doc_generator.smart_generate(
        doc_type=DocumentType.INVOICE, 
        invoice_type=InvoiceTemplate.NEON
    )
    