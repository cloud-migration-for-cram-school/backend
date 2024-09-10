from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from dotenv import load_dotenv
import os
# .envファイルから環境変数を読み込む
load_dotenv()

api_path = os.getenv('API_PATH')
folder_id = os.getenv('FOLDER_ID')

#google driveに設定
scopes = ['https://www.googleapis.com/auth/drive']

credentials = Credentials.from_service_account_file(
    api_path,
    scopes = scopes
)

class DriveService:
    def __init__(self):
        credentials = Credentials.from_service_account_file(
            api_path,
            scopes = scopes
        )
        self.drive_service = build('drive', 'v3', credentials=credentials)
        # Google Sheets APIに接続
        self.sheets_service = build('sheets', 'v4', credentials=credentials)
        self.query = f"'{folder_id}' in parents"
    
    def get_info(self):
        """
        スプレッドシートのファイル名とIDを取得
        """
        try:
            results = self.drive_service.files().list(q=self.query, fields='files(name, id)').execute()
            items = results.get('files', [])
            return items
        except Exception as e:
            print(e)