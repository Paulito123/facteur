from enum import Enum


class DocumentType(Enum):
    INVOICE = "INVOICE"
    OFFER = "OFFER"
    

class InvoiceTemplate(Enum):
    ARGENTA = "ARGENTA"
    NEON = "NEON"


class OfferTemplate(Enum):
    NEON = "NEON"


class BorderTemplate(Enum):
    NO_BORDERS = "NO_BORDERS"
    DETAIL_1 = "DETAIL_1"
