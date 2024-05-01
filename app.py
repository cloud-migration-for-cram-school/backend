from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import gspread
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
import datetime
import heapq
import json
from openpyxl.utils import get_column_letter


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

class SpreadsheetService:
    def __init__(self, credentials):
        self.credentials = credentials
        self.gc = gspread.authorize(credentials)
        #Google Drive APIに接続
        self.service = build('drive', 'v3', credentials = credentials)
        self.query = f"'{folder_id}' in parents"

    def get_info(self):
        """
        スプレッドシートのファイル名とIDを取得
        """
        try:
            results = self.service.files().list(q = self.query, fields = 'files(id, name)').execute()
            items = results.get('files', [])
            return items
        except Exception as e:
            print(e)

    def get_worksheet(self, sheetid):
        """
        sheetidのシート名とIDを辞書として返す
        """
        sheet_url = f'https://docs.google.com/spreadsheets/d/{sheetid}/edit?usp=sharing'
        spreadsheet = self.gc.open_by_url(sheet_url)
        worksheets = spreadsheet.worksheets()
        return [{'sub': ws.title, 'id': ws.id} for ws in worksheets]
    
    def get_date_info(self, id, sheetname):
        sheet_url = f'https://docs.google.com/spreadsheets/d/{id}/edit?usp=sharing'
        spreadsheet = self.gc.open_by_url(sheet_url)
        spreadsheet = spreadsheet.worksheet(sheetname)
        compar_day = spreadsheet.get_all_values(1)
        

    def get_closest_positions(self, dates_positions, target_date):
        """
        今日から最も日付が近いセルの位置を取得
        dict型 'row': 列番号, 'str_row' : '文字型の列番号(ex ABG etc)'
        """
        format_str = "%m/%d %H:%M"
        differences = [abs(datetime.datetime.strptime(target_date, format_str) - datetime.datetime.strptime(date_position[0], format_str)) for date_position in dates_positions]
        closest_indices = heapq.nsmallest(3, range(len(differences)), key=differences.__getitem__)
        return [(dates_positions[i][1]) for i in closest_indices]

    def get_old_sheet(self, postionCell, sub_name):
        for i in range(len(postionCell)):
            sheet_data = self.spreadsheet.worksheet(f'{sub_name}')
            start_row = export_cell_position(value=postionCell[i]-1)
            end_row =  export_cell_position(value=postionCell[i]+4)
            base_data = sheet_data.batch_get(f'{start_row}1:{end_row}37')
            return base_data



sp = SpreadsheetService(credentials)



#google drive内のファイル名とIDを取得
@app.get('/search')
def get_search_info():
    try:
        items = sp.get_info()
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
@app.get('/search/{sheetid}')
def user_info2(sheet_id: str):
    format_str = "%m/%d %H:%M"
    row = 2

    now_date = datetime.datetime.now()
    f_now_date = now_date.strftime(format_str)

    worksheet_lists = sp.get_worksheet(sheetid=sheet_id)
    sheet_log_dict = {}
    #シート名を取得するfor
    for worksheet in worksheet_lists:
        dates_positions = []
        #sheet_nameの中で今日から最も近い日付のセルの位置を取得するwhile
        while True:
            compar_day = worksheet.cell(1, row).value
            if compar_day is None:
                print('compar_day : None Cell')
                break #ここはリターンで返す。（シートが無い場合)
            f_format_day = datetime.datetime.strptime(compar_day, format_str).strftime(format_str)
            dates_positions.append((f_format_day, row))
            row += 7
        closest_positions = sp.get_closest_positions(dates_positions, f_now_date)
        log_sheet = sp.get_old_sheet(postionCell=closest_positions, sub_name=worksheet) #過去のセルの取得
        print(f'log_sheet : {log_sheet}')
        #sheet_log_dict[get_sheet_info.title] = log_sheet
    #return sheet_log_dict


@app.get('/search/{sheet_id}')
def user_info(sheet_id: str):
    format_str = "%m/%d %H:%M"
    now_date = datetime.datetime.now()
    f_now_date = now_date.strftime(format_str)

    worksheet_lists = sp.get_worksheet(sheetid=sheet_id)
    sheet_log_dict = {}
    for worksheet in worksheet_lists:
        dates_positions = []
        date_info = sp.get_date_info(id=sheet_id, sheetname=worksheet['sub'])



def export_cell_position(value):
    # 列のインデックスをA1表記に変換
    column_letter = get_column_letter(value)
    return column_letter




if __name__ == '__main__':
    z = user_info(sheet_id='1CEn2feUeQMfq885PVtyX96-ImOQxeqypeyePHpnPMg4')
    print(z)