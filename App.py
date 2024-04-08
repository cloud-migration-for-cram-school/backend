from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import gspread
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
import datetime

#fastapiの有効化
app = FastAPI()

# FastAPI に CORS ミドルウェアを追加
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], # 任意のオリジンからのリクエストを許可
    allow_credentials=True,
    allow_methods=["*"], # 任意の HTTP メソッドを許可
    allow_headers=["*"], # 任意のヘッダーを許可
)

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


#検索画面後の画面の遷移
@app.post('/search/result')
def user_info(username, sheetid):
    service = build('drive', 'v3', credentials = credentials)
    query = f"'{folder_id}' in parents"
    #フォルダー内のファイル名を全て取得
    results = service.files().list(q = query, fields = 'files(id, name)').execute()
    items = results.get('files', [])
    file_names = [item['name'] for item in items]
    #編集権限
    gc = gspread.authorize(credentials)
    try:
        now_date = datetime.datetime.get().now()
        #ファイル名のfor文
        for user in file_names:
            if username == user:
                sheet_url = f'https://docs.google.com/spreadsheets/d/{sheetid}/edit?usp=sharing'
                spreadsheet = gc.open_by_url(sheet_url)
                #シートを全て取得
                worksheet_lists = spreadsheet.worksheets()
                worksheet_lists.remove('閲覧用')
                #科目名のfor文
                for sheet_name in worksheet_lists:
                    get_sheet_info = spreadsheet.worksheet(f'{sheet_name}')
                    #最近更新されたセルの値を取得する 4/6ここからどうしよう 
                    #4/8 現在の時間から比較することにしようかなと思います。だけど1/31と1/1のときとか新しい生徒の情報の取得をするときに大丈夫かな
                    #シートの名前での日付の比較
                    row = 2
                    while True:
                        compar_day = get_sheet_info.cell(1, row).value
                        #日付の比較
    except:
        pass
