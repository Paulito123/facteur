from datetime import datetime, timedelta
from typing import Dict, AnyStr, Tuple
from docx import Document
from docx.shared import Cm
from enumerations import DocumentType, InvoiceTemplate, OfferTemplate, BorderTemplate
from json import load, dumps

from config import Config
from doc_helper import DocHelper


class DocProcessor:

    def __init__(self, data: Dict = {}, is_test_run: bool = True) -> None:
        
        self.db = None
        self.set_data(data)
        self.is_test_run = is_test_run

        try:
            with open(Config.PATH_DB, "r") as f:
                self.db = load(f)
        except Exception as e:
            raise Exception(f"Could not initialize DB: {e}")
    

    def set_data(self, data: Dict) -> None:
        self.data = data
        

    def __check_data(self, data) -> bool:
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
    

    def __generate(self, data: Dict, doc_name: AnyStr) -> None:
        """
        fucntion to generate the actual document
        """
        # Check if the data is valid
        # if not self.__check_data(data):
        #     return None
        
        # create a new document
        dhelpr = DocHelper()
        document = Document()

        # set the styles for the document
        dhelpr.set_styles(document)

        # set the header, body and footer
        dhelpr.set_header(document, data["header"])
        dhelpr.set_body(document, data["body"])
        dhelpr.set_footer(document, data["footer"])

        # save the document and convert it to pdf
        document.save(f'files/invoices/{doc_name}.docx')
        dhelpr.convert_to_pdf(f'files/invoices/{doc_name}.docx', f'files/invoices/{doc_name}.pdf')

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
                       invoice_type: InvoiceTemplate = InvoiceTemplate.NEON,
                       is_test_run: bool = True) -> None:
        
        self.is_test_run = is_test_run
        
        # Check if the data is valid
        # if not self.__check_data(self.data):
        #     return None
        
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
                doc_data["header"]["due_date"] = (datetime.now() + timedelta(days=30)).strftime("%d-%m-%Y")
            
            doc_data["body"]["due_date"] = doc_data["header"]["due_date"]
            
            # debtor
            debtor = self.db["companies"][self.data["debtor_id"]]
            for key in debtor.keys():
                doc_data["header"]['debtor_' + key] = debtor[key]

            # policies
            policies = [self.db["policies"][pid] for pid in self.db["defaults"]["policy_ids"]]
            doc_data["body"]["policies"] = policies

            # symbol
            currency = self.db["currencies"][self.db["defaults"]["currency_id"]]
            doc_data["body"]["symbol"] = currency["symbol"]

            # payment date
            if "payment_date" in self.data:
                doc_data["body"]["payment_date"] = self.data["payment_date"]

            # items
            invoice_base_amt = 0
            invoice_vat_amt = 0
            invoice_total_amt = 0
            doc_data["body"]["items"] = {}

            if invoice_type == InvoiceTemplate.NEON:
                items = self.data["items"]

                for item_key in items.keys():
                    # calculations
                    price = items[item_key]["price"]
                    base_amt = items[item_key]["qty"] * price
                    vat_amt = round(base_amt * items[item_key]["vat_pct"], 2)
                    total_amt = base_amt + vat_amt
                    invoice_base_amt += base_amt
                    invoice_vat_amt += vat_amt
                    invoice_total_amt += total_amt

                    # building items
                    doc_data["body"]["items"][item_key] = {}
                    doc_data["body"]["items"][item_key]["description"] = items[item_key]["description"]
                    doc_data["body"]["items"][item_key]["qty"] = items[item_key]["qty"]
                    doc_data["body"]["items"][item_key]["unit_amt"] = price
                    doc_data["body"]["items"][item_key]["base_amt"] = base_amt
                    doc_data["body"]["items"][item_key]["vat_amt"] = vat_amt
                    doc_data["body"]["items"][item_key]["total_amt"] = total_amt
            
            elif invoice_type == InvoiceTemplate.ARGENTA:
                consultancy_days = self.data["consultancy_days"]

                # day rate
                if 'day_rate' in self.data:
                    day_rate = self.data["day_rate"]
                else:
                    day_rate = self.db["defaults"]["argenta"]["day_rate"]

                base_amt = consultancy_days * day_rate
                vat_amt = round(base_amt * 0.21, 2)

                doc_data["body"]["items"]["1"] = {
                    "description": self.db["defaults"]["argenta"]["item_description"],
                    "qty": consultancy_days,
                    "unit_amt": day_rate,
                    "base_amt": base_amt,
                    "vat_amt": vat_amt,
                    "total_amt": base_amt + vat_amt
                }

                invoice_base_amt = base_amt
                invoice_vat_amt = vat_amt
                invoice_total_amt = base_amt + vat_amt

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

                    # add creditor bank account for trailing message
                    if key == "bank_account":
                        doc_data["body"]['creditor_' + key] = creditor[key]

                else:
                    # invoice number
                    next_seq = self.get_next_doc_sequence(doc_type, creditor['last_sequences']["invoice"])
                    doc_data["header"][next_seq[0]] = next_seq[1]
            
            self.generate_invoice(doc_data)

        elif doc_type == DocumentType.OFFER:
            doc_data["header"]["title"] = "Offerte"
            self.generate_offer(doc_data)
        

    def generate_offer(self, data: Dict):
        
        # Create a offer
        doc_name = f'O_{data["header"]["offer_nr"]}'
        self.__generate(self.data, doc_name)


    def generate_invoice(self, data: Dict):
        
        # Create an invoice
        doc_name = f'I_{data["header"]["invoice_nr"]}'
        self.__generate(data, doc_name)

        if not self.is_test_run:
            # Update the database
            self.save_db(DocumentType.INVOICE, increase_seq=True)


    def save_db(self, DocType: DocumentType, increase_seq: bool = False) -> None:
        """
        Save the current state of the database to a file.
        Args:
            DocType (DocumentType): The type of document being processed (e.g., OFFER or INVOICE).
            increase_seq (bool, optional): Flag indicating whether to increment the sequence number for the document type. Defaults to False.
        Returns:
            None
        """
        
        seq_id = "offer" if DocType == DocumentType.OFFER else "invoice"
        
        if increase_seq:
            self.db["companies"][self.data["creditor_id"]]["last_sequences"][seq_id] += 1
        
        with open(Config.PATH_DB, "w") as f:
            f.write(dumps(self.db, sort_keys=True, indent=4))
    