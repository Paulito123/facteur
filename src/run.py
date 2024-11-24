import json
from time import sleep

from docx import Document
from doc_helper import DocHelper
from doc_processor import DocProcessor
from enumerations import DocumentType, InvoiceTemplate

from config import Config
from gdrive import GDrive
from gmail import Gmail


if __name__ == "__main__":
    ############################################
    # Example usage of the DocHelper class
    ############################################

    # data = {}
    
    # with open("files/config/prd_arg_202411.json", "r") as f:
    #     data = json.load(f)

    # doc_generator = DocProcessor(data)
    # doc_generator.smart_generate(
    #     DocumentType.INVOICE, 
    #     invoice_type=InvoiceTemplate.ARGENTA, 
    #     is_test_run=True
    # )

    # with open("files/config/prd_isc.json", "r") as f:
    #     data = json.load(f)
    
    # doc_generator = DocProcessor(data)
    # doc_generator.smart_generate(
    #     DocumentType.INVOICE,
    #     is_test_run=True
    # )

    ############################################
    # Example usage of Google API classes
    ############################################
    token_file = Config.PATH_CONFIG + Config.GTOKEN_FILE_NAME
    secret_file = Config.PATH_CONFIG + Config.CLIENT_TOKEN

    ############################################
    # Example usage of the GDrive class
    ############################################
    # gd = GDrive(secret_file, token_file)

    # # upload a file
    # gd.upload_file(parent_folder_id=Config.DIR_ID_ARGENTA, file_path='files/invoices/voorbeeld2.pdf')
    
    # Create a new folder
    # p.create_folder("MyNewFolder")
    
    # List folders and files
    # p.list_folder(parent_folder_id=Config.DIR_ID_ARGENTA)

    # Delete a file or folder by ID
    # p.delete_files("your_file_or_folder_id")

    # Download a file by its ID
    # p.download_file(Config.DIR_ID_ARGENTA, "files/some_invoice.pdf")

    ############################################
    # Example usage of the Gmail class
    ############################################
    gm = Gmail(secret_file, token_file)

    # print("create draft")
    # draft = gm.create_draft(
    #     to="subhero007@hotmail.com",
    #     subject="Test draft",
    #     message_text="Hello, this is a test draft message."
    # )

    # print("sleep 15 secs")
    # sleep(15)

    # print("send draft")
    # gm.send_draft(draft["id"])

    # print("sleep 15 secs")
    # sleep(15)
    
    print("send direct message with attachment")
    gm.send_message(
        to="subhero007@hotmail.com",
        subject="Test message",
        message_text="Hello, this is a test message with an attachment.",
        attachments=["files/invoices/example.pdf"]
    )
