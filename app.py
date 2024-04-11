from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import gspread
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
import datetime
import heapq
import json

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
@app.post('/search/{sheetid}')
def user_info(sheetid: str):
    format_str = "%m/%d %H:%M"
    row = 2
    gc = gspread.authorize(credentials)
    
    now_date = datetime.datetime.now()
    f_now_date = now_date.strftime(format_str)
    sheet_url = f'https://docs.google.com/spreadsheets/d/{sheetid}/edit?usp=sharing'
    spreadsheet = gc.open_by_url(sheet_url)
    worksheet_lists = spreadsheet.worksheets()
    sheet_log_dict = {}
    #シート名を取得するfor
    for worksheet in worksheet_lists:
        get_sheet_info = worksheet
        dates_positions = []
        #sheet_nameの中で今日から最も近い日付のセルの位置を取得するwhile
        while True:
            compar_day = get_sheet_info.cell(1, row).value
            if compar_day is None:
                print('None Cell')
                break #ここはリターンで返す。（シートが無い場合)
            f_format_day = datetime.datetime.strptime(compar_day, format_str).strftime(format_str)
            dates_positions.append((f_format_day, row))
            row += 7
        closest_positions = get_closest_positions(dates_positions, f_now_date)
        log_sheet = get_old_sheet(postionCell=closest_positions, sheet_data=get_sheet_info) #過去のセルの取得
        sheet_log_dict[get_sheet_info.title] = log_sheet
    return sheet_log_dict


#取得した日付のセルの位置を取得
def get_closest_positions(dates_positions, target_date):
    format_str = "%m/%d %H:%M"
    differences = [abs(datetime.datetime.strptime(target_date, format_str) - datetime.datetime.strptime(date_position[0], format_str)) for date_position in dates_positions]
    #nsmallestの引数(4, ..)で取得する枚数を指定する
    closest_indices = heapq.nsmallest(4, range(len(differences)), key=differences.__getitem__)
    return [dates_positions[i][1] for i in closest_indices]

#過去の資料の取得
def get_old_sheet(postionCell, sheet_data):
    old_sheets = []  # 複数のシートデータを格納するリスト
    for i in range(len(postionCell)):
        #log_sheet_dataは報告書と同じ形になるようにしています。多少見づらいですが直感的に分かりやすくするためにはしょうがない犠牲です。 
        log_sheet_data = {
            "basicInfo": {
                "dateAndTime": sheet_data.cell(1, postionCell[i]).value, "subjectName": sheet_data.cell(1, postionCell[i]+1).value, "teacherName": sheet_data.cell(1, postionCell[i]+2).value,
                "progressInSchool": sheet_data.cell(2, postionCell[i]).value, "homeworkProgress": sheet_data.cell(2, postionCell[i]+1).value, "homeworkAccuracy": sheet_data.cell(2, postionCell[i]+2).value,
            }
        }
        old_sheets.append(log_sheet_data)
    # old_sheets を JSON 形式の文字列に変換して返す
    return json.dumps(old_sheets, ensure_ascii=False, indent=4)

if __name__ == '__main__':
    pass
    #print(user_info(sheetid='1CEn2feUeQMfq885PVtyX96-ImOQxeqypeyePHpnPMg4'))