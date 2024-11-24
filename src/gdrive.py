import os.path
import io
from typing import AnyStr, List

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

from config import Config


class GDrive():

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
            self.service = build("drive", "v3", credentials=self.creds)

        except HttpError as error:
            # TODO(developer) - Handle errors from drive API.
            print(f"An error occurred: {error}")

        
    def create_folder(self, folder_name, parent_folder_id=None):
        """
        Creates a folder in Google Drive.
        Args:
            folder_name (str): The name of the folder to create.
            parent_folder_id (str, optional): The ID of the parent folder where the new folder will be created. 
                                              If None, the folder will be created in the root directory.
        Returns:
            str: The ID of the created folder.
        Raises:
            googleapiclient.errors.HttpError: If the request to the Google Drive API fails.
        """

        # folder metadata
        folder_metadata = {
            'name': folder_name,
            "mimeType": "application/vnd.google-apps.folder",
            'parents': [parent_folder_id] if parent_folder_id else []
        }

        # create a folder in Google Drive
        created_folder = self.service.files().create(
            body=folder_metadata,
            fields='id'
        ).execute()

        # print the ID of the created folder
        print(f'Created Folder ID: {created_folder["id"]}')

        # return the ID of the created folder
        return created_folder["id"]


    def list_folder(self, parent_folder_id=None, delete=False) -> List[dict]:
        """
        Lists the contents of a specified Google Drive folder.
        Args:
            parent_folder_id (str, optional): The ID of the parent folder to list contents from. 
                                              If None, lists from the root directory. Defaults to None.
            delete (bool, optional): If True, deletes the files after listing them. Defaults to False.
        Returns:
            List[dict]: A list of dictionaries containing file information such as 'id', 'name', and 'mimeType'.
        """
        
        
        # list the contents of the specified folder
        results = self.service.files().list(
            q=f"'{parent_folder_id}' in parents and trashed=false" if parent_folder_id else None,
            pageSize=1000,
            fields="nextPageToken, files(id, name, mimeType)"
        ).execute()

        # get the items from the results
        items = results.get('files', [])

        return items

        # # print the items in the folder
        # if not items:
        #     print("No folders or files found in Google Drive.")
        # else:
        #     print("Folders and files in Google Drive:")
        #     for item in items:
        #         if item['mimeType'] == 'application/vnd.google-apps.folder':
        #             print(f"Directory: {item['name']}, ID: {item['id']}, Type: {item['mimeType']}")
        #         else:
        #             print(f"File: {item['name']}, ID: {item['id']}, Type: {item['mimeType']}")

                # if delete:
                #     self.delete_files(item['id'])


    def delete_files(self, file_or_folder_id):
        """
        Deletes a file or folder from Google Drive using its ID.
        Args:
            file_or_folder_id (str): The ID of the file or folder to be deleted.
        Returns:
            None
        Raises:
            Exception: If there is an error during the deletion process, it will be caught and printed.
        """
        
        try:
            # delete the file or folder
            self.service.files().delete(fileId=file_or_folder_id).execute()
            print(f"Successfully deleted file/folder with ID: {file_or_folder_id}")
        except Exception as e:
            print(f"Error deleting file/folder with ID: {file_or_folder_id}")
            print(f"Error details: {str(e)}")


    def download_file(self, file_id, destination_path):
        """
        Downloads a file from Google Drive to a specified local destination.
        Args:
            file_id (str): The ID of the file to be downloaded from Google Drive.
            destination_path (str): The local file path where the downloaded file will be saved.
        Returns:
            None
        Raises:
            googleapiclient.errors.HttpError: If an error occurs during the download process.
        """
        
        # request the file from Google Drive
        request = self.service.files().get_media(fileId=file_id)
        fh = io.FileIO(destination_path, mode='wb')
        
        # create a file downloader
        downloader = MediaIoBaseDownload(fh, request)
        
        # download the file in chunks
        done = False
        while not done:
            status, done = downloader.next_chunk()
            print(f"Download {int(status.progress() * 100)}%.")
    

    def upload_file(self, file_path, parent_folder_id=None):
        """
        Uploads a file to Google Drive.
        Args:
            file_path (str): The path to the file to be uploaded.
            parent_folder_id (str, optional): The ID of the parent folder in Google Drive where the file will be uploaded. 
                                              If not provided, the file will be uploaded to the root directory.
        Returns:
            str: The ID of the uploaded file.
        Raises:
            googleapiclient.errors.HttpError: If an error occurs during the file upload.
        Example:
            >>> gdrive.upload_file('/path/to/file.pdf', 'parent_folder_id')
            Uploaded File ID: 1ZdR3L4f5g6H7i8J9k0L
        """
        
        # file metadata
        file_metadata = {
            'name': os.path.basename(file_path),
            'parents': [parent_folder_id] if parent_folder_id else []
        }

        # create an uploader for the file
        media = MediaFileUpload(file_path, mimetype='application/pdf')
        
        # upload the file to Google Drive
        uploaded_file = self.service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()
        
        print(f'Uploaded File ID: {uploaded_file["id"]}')
        return uploaded_file["id"]
