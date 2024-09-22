from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
import os
from dotenv import load_dotenv

load_dotenv()
API_PATH = os.getenv('API_PATH')
SCOPES = ['https://www.googleapis.com/auth/drive']

class SpreadsheetServiceInit:
    def __init__(self, fileID=None):
        self.credentials = Credentials.from_service_account_file(
            API_PATH,
            scopes=SCOPES
        )
        self.fileID = fileID
        self.service = build('sheets', 'v4', credentials=self.credentials)
