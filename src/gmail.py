import os.path
import base64
import mimetypes
from typing import AnyStr, Dict, List, Union

from email.message import EmailMessage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from config import Config


class Gmail:

    def __init__(self, secret_file: AnyStr, token_file: AnyStr):
        """
        Initializes the Google Drive service with the provided credentials.
        Args:
            secret_file (AnyStr): Path to the client secrets file.
            token_file (AnyStr): Path to the token file where credentials are stored.
        Attributes:
            creds (Credentials): The authenticated user's credentials.
            service (Resource): The Google Drive service instance.
        Raises:
            HttpError: If an error occurs while building the Google Drive service.
        """

        self.creds = None
        self.service = None

        # Load the credentials from the token file if it exists
        if os.path.exists(token_file):
            self.creds = Credentials.from_authorized_user_file(token_file, Config.APP_SCOPES)
        
        # If there are no (valid) credentials available, let the user log in.
        if not self.creds or not self.creds.valid:

            # Refresh the credentials if they are expired
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    secret_file, Config.APP_SCOPES
                )
                self.creds = flow.run_local_server(port=0)
            
            # Save the credentials for the next run
            with open(token_file, "w") as token:
                token.write(self.creds.to_json())
        
        try:

            # Build the Google Drive service
            self.service = build("gmail", "v1", credentials=self.creds)

        except HttpError as error:
            # TODO(developer) - Handle errors from drive API.
            print(f"An error occurred: {error}")
    

    def __build_file_part(self, file):
        """Creates a MIME part for a file.

        Args:
            file: The path to the file to be attached.

        Returns:
            A MIME part that can be attached to a message.
        """
        content_type, encoding = mimetypes.guess_type(file)

        print(f"file: {file}")
        print(f"content_type: {content_type}")
        print(f"encoding: {encoding}")

        if content_type is None or encoding is not None:
            content_type = "application/octet-stream"
        main_type, sub_type = content_type.split("/", 1)
        if main_type == "text":
            with open(file, "rb"):
                msg = MIMEText("r", _subtype=sub_type)
        elif main_type == "image":
            with open(file, "rb"):
                msg = MIMEImage("r", _subtype=sub_type)
        elif main_type == "audio":
            with open(file, "rb"):
                msg = MIMEAudio("r", _subtype=sub_type)
        else:
            with open(file, "rb") as f:
                msg = MIMEBase(main_type, sub_type)
                msg.set_payload(f.read())
        filename = os.path.basename(file)
        msg.add_header("Content-Disposition", "attachment", filename=filename)
        return msg
    

    def __compose_message(self, to: AnyStr, subject: AnyStr = "", message_text: AnyStr = "", attachments: List[AnyStr]  = None) -> Union[None, Dict[AnyStr, Dict[AnyStr, AnyStr]]]:
        """
        Composes an email message with optional attachments.
        Args:
            to (AnyStr): The recipient's email address.
            subject (AnyStr, optional): The subject of the email. Defaults to an empty string.
            message_text (AnyStr, optional): The body text of the email. Defaults to an empty string.
            attachments (List[AnyStr], optional): A list of file paths to attach to the email. Defaults to None.
        Returns:
            Union[None, Dict[AnyStr, Dict[AnyStr, AnyStr]]]: A dictionary containing the encoded email message if successful, 
            otherwise None.
        Raises:
            Exception: If an error occurs during the composition of the email.
        """
        
        try:
            if attachments:
                message = MIMEMultipart()

                message.attach(MIMEText(message_text, 'plain'))

                for attachment in attachments:
                    # guessing the MIME type
                    type_subtype, _ = mimetypes.guess_type(attachment)
                    maintype, subtype = type_subtype.split("/")

                    attachment_data = self.__build_file_part(attachment)

                    message.add_attachment(attachment_data, maintype, subtype)
            else:
                message = EmailMessage()

                message.set_content(message_text)

            message["To"] = to
            message["From"] = Config.FROM_ADDRESS
            message["Subject"] = subject

            # encoded message
            encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
            
            # return the message
            return {"message": {"raw": encoded_message}}

        except Exception as e:
            print(f"An error occurred: {e}")
            
        return None


    def create_draft(self, to: AnyStr, subject: AnyStr, message_text: AnyStr, attachments: List[AnyStr] = None) -> dict:
        """
        Creates a draft email with the specified recipient, subject, and message text.
        Args:
            to (AnyStr): The recipient's email address.
            subject (AnyStr): The subject of the email.
            message_text (AnyStr): The body text of the email.
        Returns:
            dict: The created draft's metadata, including its ID and message content, or None if an error occurred.
        Raises:
            HttpError: If an error occurs while creating the draft.
        """

        try:
            body = self.__compose_message(to, subject, message_text, attachments)

            # pylint: disable=E1101
            draft = (
                self.service.users()
                .drafts()
                .create(userId="me", body=body)
                .execute()
            )
            print(f'Draft id: {draft["id"]}\nDraft message: {draft["message"]}')
        except HttpError as error:
            print(f"An error occurred: {error}")
            draft = None
        return draft


    def send_message(self, to: AnyStr, subject: AnyStr, message_text: AnyStr, attachments: List[AnyStr] = None) -> dict:
        """Create and send an email message
        Print the returned  message id
        Returns: Message object, including message id

        Load pre-authorized user credentials from the environment.
        TODO(developer) - See https://developers.google.com/identity
        for guides on implementing OAuth2 for the application.
        """

        try:
            body = self.__compose_message(to, subject, message_text, attachments)

            # pylint: disable=E1101
            send_message = (
                self.service.users()
                .messages()
                .send(userId="me", body=body)
                .execute()
            )
            print(f'Message Id: {send_message["id"]}')
        except HttpError as error:
            print(f"An error occurred: {error}")
            send_message = None
        return send_message


    def send_draft(self, draft_id: AnyStr) -> dict:
        """
        Sends a precomposed draft email by its ID.
        Args:
            draft_id (AnyStr): The ID of the draft to be sent.
        Returns:
            dict: The sent message's metadata, including its ID, or None if an error occurred.
        Raises:
            HttpError: If an error occurs while sending the draft.
        """
        try:
            # pylint: disable=E1101
            send_draft = (
                self.service.users()
                .drafts()
                .send(userId="me", body={"id": draft_id})
                .execute()
            )
            print(f'Message Id: {send_draft["id"]}')
        except HttpError as error:
            print(f"An error occurred: {error}")
            send_draft = None
        return send_draft


    def bla(self, string: str) -> dict:
        return self.__build_file_part(string)


if __name__ == "__main__":
    token_file = Config.PATH_CONFIG + Config.GTOKEN_FILE_NAME
    secret_file = Config.PATH_CONFIG + Config.CLIENT_TOKEN

    gm = Gmail(secret_file, token_file)

    message = gm.bla("files/invoices/example.pdf")

    # print(message)