from typing import Dict, AnyStr
from docx import Document
from docx.shared import Cm
import enumerations


class DocGenerator:
    def __init__(self):
        pass

    def __check_data(self, data) -> bool:
        """
        Validates the provided data dictionary to ensure it contains the required fields.
        Args:
            data (dict): The data dictionary to validate. It must contain the following keys:
                - "creditor_id"
                - "debtor_id"
                - "debtor_person_id"
                - "invoice_date"
                - "due_date"
                - "items" (list of dictionaries, each containing the keys "description", "qty", "vat_pct", and "base_amt")
        Returns:
            bool: True if the data contains all required fields, False otherwise.
        Raises:
            Exception: If any of the required fields are missing in the items list.
        """
        try:
            # level 1 check
            for s in ["creditor_id", "debtor_id", "debtor_person_id", "invoice_date", "due_date", "items"]:
                if s not in data:
                    raise Exception(f"Missing field: {s}")

            # level 2 check
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

    def generate_by_id(self, id: AnyStr):
        pass

    def generate_argenta_invoice(self, data: Dict):
        
        # Generate the invoice
        return self.__generate(data)
        
    def generate_neon_invoice(self, data: Dict):
        
        # Generate the invoice
        return self.__generate(data)
        