from googleapiclient.errors import HttpError
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
import json
import service.transform_data
from service.spreadsheet_service import SpreadsheetService
from service.drive_service import DriveService
from fastapi import Request

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

#google drive内のファイル名とIDを取得
@app.get('/search')
def get_drive_file_info():
    deiveservice = DriveService()
    try:
        items = deiveservice.get_info()
        # Create a list to hold the file names and IDs
        file_info = []
        # Loop through the items and add the file name and ID to the list
        for item in items:
            file_info.append({'label': item['name'], 'value': item['id']})
        return file_info
    except HttpError as error:
        message = f'エラー: {error.content.decode("utf-8")}'
        return message

@app.get('/search/subjects/{sheet_id}')
def get_subject_list(sheet_id: str):
    sp = SpreadsheetService(fileID=sheet_id)
    subject = sp.get_worksheets()
    return subject

# 一旦無効化 (多分使わない気がする)
def get_user_info(sheet_id: str, subjects_id: int):
    sp = SpreadsheetService(fileID=sheet_id)
    # 指定されたシートIDに対応するシートの情報を取得
    subjects = sp.get_worksheets()
    # subjects_idを使用して、対応するシートの名前を検索
    for subject in subjects:
        if subject['value'] == subjects_id:
            sheet_name = subject['label']
            break
    else:
        return {"error": "Subject not found"}
    date_info = sp.find_closest_dates(worksheetname=sheet_name, start_row=2)
    # date_infoを大きい順に並び替える
    sorted_date_info = sorted(date_info, key=lambda x: x['row'], reverse=True)

    positions = [item['row'] for item in sorted_date_info]
    old_sheet = sp.get_old_sheet_data(postionCell=positions[0], worksheetname=sheet_name)

    transformed_data = service.transform_data.transform_data(old_sheet[0])
    return json.dumps(transformed_data, ensure_ascii=False, indent=2)

@app.get('/search/subjects/reports/{sheet_id}/{subjects_id}')
def get_report(sheet_id: str, subjects_id: str):
    try:
        sp = SpreadsheetService(fileID=sheet_id)

        # シート内で取得する位置を探索
        date_info = sp.find_exponential_dates(subject_id=int(subjects_id), expotent_base=7, start_row=2)
        if date_info[0]['position'] != 0:
            # 過去の報告書が存在する場合の処理
            meticulous_date_info = sp.find_exponential_dates(subject_id=int(subjects_id), expotent_base=7, start_row=date_info[-1]['position'])
            liner_search = sp.find_closest_dates(subject_id=int(subjects_id), start_row=meticulous_date_info[-1]['position'])

            sorted_date_info = sorted(liner_search, key=lambda x: x['row'], reverse=True)
            positions = [item['row'] for item in sorted_date_info]

            # データ取得とデータの整形
            old_sheet = sp.get_old_sheet_data(postionCell=positions[0], subject_id=int(subjects_id))
            transformed_data = service.transform_data.transform_data(old_sheet[0])
            return transformed_data  # JSON形式としてそのまま返す
        else: # 過去の報告書が存在しない場合
            return service.transform_data.initialize_mapping_with_defaults()
    except Exception as e:
        print(f"エラーが発生しました(get_report): {e}")
        return {"error": f"エラーが発生しました: {str(e)}"}


@app.post('/submit/report/{sheet_id}/{subjects_id}')
async def submit_report(sheet_id: str, subjects_id: str, request: Request):
    try:
        # JSON データを取得
        report_data = await request.json()  # awaitを使って結果を待つ
        # SpreadsheetService のインスタンスを作成
        sp = SpreadsheetService(fileID=sheet_id)
       
        # シート内で入力する位置を探索
        date_info = sp.find_exponential_dates(subject_id=int(subjects_id), expotent_base=7, start_row=2)
        if date_info[0]['position'] != 0:
            # 過去の報告書が存在する場合
            meticulous_date_info = sp.find_exponential_dates(subject_id=int(subjects_id), expotent_base=7, start_row=date_info[-1]['position'])
            liner_search = sp.find_closest_dates(subject_id=int(subjects_id), start_row=meticulous_date_info[-1]['position'])

            # 最も左側の空セルの位置を決定
            target_position = max(item['row'] for item in liner_search) + 6
        else:
            # 過去の報告書が存在しない場合
            # 最初の報告書に入力する最初の位置は2列目固定
            target_position = 1
        
        # JSONデータをスプレッドシート形式に変換
        transformed_data = service.transform_data.reverse_transform_data(report_data)
        # データをスプレッドシートに登録
        sp.update_report(target_position, transformed_data, subject_id=int(subjects_id))
        return {"status": "success", "message": "Report submitted successfully."}
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        return {"error": f"エラーが発生しました: {str(e)}"}

@app.post('/submit/report/old/{sheet_id}/{subjects_id}')
async def submit_report_old(sheet_id: str, subjects_id: str, request: Request):
    try:
        # JSON データを取得
        report_data = await request.json()  # awaitを使って結果を待つ
        # SpreadsheetService のインスタンスを作成
        sp = SpreadsheetService(fileID=sheet_id)
        
        # シート内で入力する位置を探索
        date_info = sp.find_exponential_dates(subject_id=int(subjects_id), expotent_base=7, start_row=2)
        meticulous_date_info = sp.find_exponential_dates(subject_id=int(subjects_id), expotent_base=7, start_row=date_info[-1]['position'])
        liner_search = sp.find_closest_dates(subject_id=int(subjects_id), start_row=meticulous_date_info[-1]['position'])

        # 最新の過去のデータの位置を取得
        target_position = max(item['row'] for item in liner_search) - 1
        # JSONデータをスプレッドシート形式に変換
        transformed_data = service.transform_data.reverse_transform_data(report_data)
        # データをスプレッドシートに登録
        sp.update_report(target_position, transformed_data, subject_id=int(subjects_id))
        return {"status": "success", "message": "Report submitted successfully."}
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        return {"error": f"エラーが発生しました: {str(e)}"}

if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)