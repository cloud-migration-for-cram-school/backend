from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from fastapi import FastAPI


#fastapiの有効化
app = FastAPI()

api_path = r'C:\Users\sabax\Documents\Python\API_check\apitest-418907-1658bd7509d0.json'
folder_id = '163LOPPPAgKUZXpegrYbBbZkGntFnH5-v'

#google driveに設定
scopes = ['https://www.googleapis.com/auth/drive']

credentials = Credentials.from_service_account_file(
    api_path,
    scopes = scopes
)

#google drive内のファイル名とIDを取得
@app.get('/search')
def get_search_info():
    #Google Drive APIに接続
    service = build('drive', 'v3', credentials = credentials)
    query = f"'{folder_id}' in parents"
    try:
        results = service.files().list(q = query, fields = 'files(id, name)').execute()
        items = results.get('files', [])
        # Create a list to hold the file names and IDs
        file_info = []
        # Loop through the items and add the file name and ID to the list
        for item in items:
            file_info.append({'name': item['name'], 'id': item['id']})
        return file_info
    except HttpError as error:
        message = f'エラー: {error.content.decode("utf-8")}'
        return message
