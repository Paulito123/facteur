import datetime

from typing import Dict, AnyStr, Tuple
from docx import Document
from docx.shared import Cm
from enumerations import DocumentType, InvoiceTemplate, OfferTemplate, BorderTemplate
from json import load

from config import Config
from doc_helper import DocHelper


class DocProcessor:

    def __init__(self, data: Dict = {}):
        self.db = None
        self.set_data(data)

        try:
            with open(Config.PATH_DB, "r") as f:
                self.db = load(f)
        except Exception as e:
            Exception(f"Could not initialize DB: {e}")
    

    def set_data(self, data: Dict) -> None:
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


    def get_next_doc_sequence(self, doc_type: DocumentType, last_seq: int) -> Tuple[AnyStr, AnyStr]:
        """
        Function to get the next document sequence number for the given document type.
        Args:
            doc_type (DocumentType): The document type.
            last_seq (Last sequence): The last sequence number used for the document type.
        Returns:
            Tuple: [the field identifier, the doc sequence number]
        """
        doc_number = f'{datetime.now().year}-{last_seq + 1}'
        field_id = 'invoice_nr' if doc_type == DocumentType.INVOICE else 'offer_nr'
        
        return (field_id, doc_number)
        

    def smart_generate(self,
                       doc_type: DocumentType = DocumentType.INVOICE, 
                       invoice_type: InvoiceTemplate = InvoiceTemplate.NEON) -> None:
        
        # Check if the data is valid
        if not self.__check_data(self.data):
            return None
        
        doc_data = {
            "header": {
                "path_image": 'files/images/tokaio.png'
            },
            "body": {},
            "footer": {}
        }
        
        # all the default selection logic goes here
        if doc_type == DocumentType.INVOICE:
            
            # title
            doc_data["header"]["title"] = "Factuur"
            
            # invoice date
            if "invoice_date" in self.data:
                doc_data["header"]["invoice_date"] = self.data["invoice_date"]
            else:
                doc_data["header"]["invoice_date"] = datetime.now().strftime("%d-%m-%Y")
            
            # delivery date
            if "delivery_date" in self.data:
                doc_data["header"]["delivery_date"] = self.data["delivery_date"]
            else:
                doc_data["header"]["delivery_date"] = datetime.now().strftime("%d-%m-%Y")
            
            # due date
            if "due_date" in self.data:
                doc_data["header"]["due_date"] = self.data["due_date"]
            else:
                doc_data["header"]["due_date"] = (datetime.now() + datetime.timedelta(days=10)).strftime("%d-%m-%Y")
            
            # debtor
            debtor = self.db["companies"][self.data["debtor_id"]]
            for key in debtor.keys():
                doc_data["header"]['debtor_' + key] = debtor[key]

            # policies
            policies = [self.db["policies"][pid] for pid in self.db["defaults"]["policy_ids"]]
            doc_data["body"]["policies"] = policies

            # symbol
            currency = self.db["currencies"][self.data["defaults"]["currency_id"]]
            doc_data["body"]["symbol"] = currency["symbol"]

            # items
            invoice_base_amt = 0
            invoice_vat_amt = 0
            invoice_total_amt = 0
            doc_data["body"]["items"] = {}

            if invoice_type == InvoiceTemplate.NEON:
                items = self.data["items"]

                for item_key in items.keys():
                    doc_data["body"]["items"][item_key] = item_key
                    doc_data["body"]["items"][item_key]["description"] = items[item_key]["description"]
                    doc_data["body"]["items"][item_key]["qty"] = items[item_key]["qty"]
                    doc_data["body"]["items"][item_key]["base_amt"] = items[item_key]["price"] * items[item_key]["price"]
                    doc_data["body"]["items"][item_key]["vat_amt"] = doc_data["body"]["items"]["base_amt"] * items[item_key]["vat_pct"]
                    doc_data["body"]["items"][item_key]["total_amt"] = doc_data["body"]["items"]["base_amt"] + doc_data["body"]["items"]["vat_amt"]
                    invoice_base_amt += doc_data["body"]["items"][item_key]["base_amt"]
                    invoice_vat_amt += doc_data["body"]["items"][item_key]["vat_amt"]
                    invoice_total_amt += (invoice_base_amt + invoice_vat_amt)
            
            elif invoice_type == InvoiceTemplate.ARGENTA:
                consultancy_days = self.data["consultancy_days"]
                base_amt = consultancy_days * self.db["defaults"]["argenta"]["day_rate"]
                doc_data["body"]["items"]["1"] = {
                    "description": self.db["defaults"]["argenta"]["item_description"],
                    "qty": consultancy_days,
                    "base_amt": base_amt,
                    "vat_amt": base_amt * 0.21,
                    "total_amt": base_amt * 1.21
                }
                invoice_base_amt = base_amt
                invoice_vat_amt = doc_data["body"]["items"]["1"]["vat_amt"]
                invoice_total_amt = doc_data["body"]["items"]["1"]["total_amt"]
            else:
                raise Exception("Invalid invoice type")
            
            # totals
            doc_data["body"]["invoice_base_amt"] = invoice_base_amt
            doc_data["body"]["invoice_vat_amt"] = invoice_vat_amt
            doc_data["body"]["invoice_total_amt"] = invoice_total_amt

            # creditor
            if 'creditor_id' in self.data:
                creditor_id = self.data['creditor_id']
            else:
                creditor_id = self.db["defaults"]["creditor_id"]

            creditor = self.db["companies"][creditor_id]
            for key in creditor.keys():
                if key not in ["last_sequences"]:
                    doc_data["footer"]['creditor_' + key] = creditor[key]
                else:
                    # invoice number
                    next_seq = self.get_next_doc_sequence(doc_type, creditor['last_sequences']["invoice"])
                    doc_data["header"][next_seq[0]] = next_seq[1]

        elif doc_type == DocumentType.OFFER:
            doc_data["header"]["title"] = "Offerte"
            self.generate_offer(self.data)
        
        self.generate_invoice(self.data)


    def generate_offer(self):
        
        # Create a 
        self.__generate(self.data)

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
    