from dotenv import dotenv_values 


envs = dotenv_values(".env")

class Config:
    PATH_DB = envs["PATH_DB"]
    PATH_OUT = envs["PATH_OUT"]
    PATH_CONFIG = envs["PATH_CONFIG"]

    CLIENT_ID = envs["CLIENT_ID"]
    CLIENT_TOKEN = envs["CLIENT_TOKEN"]
    GTOKEN_FILE_NAME = envs["GTOKEN_FILE_NAME"]

    DIR_ID_ARGENTA = envs["DIR_ID_ARGENTA"]

    # If modifying these scopes, delete the file token.json 
    # abd clear browser cache if issues persist. 
    APP_SCOPES = [
        # "https://www.googleapis.com/auth/drive.metadata.readonly",
        "https://www.googleapis.com/auth/drive.file",
        # "https://www.googleapis.com/auth/drive",
        # "https://www.googleapis.com/auth/drive.appdata",
        # "https://www.googleapis.com/auth/drive.appfolder",
        # "https://www.googleapis.com/auth/drive.install",
    ]
