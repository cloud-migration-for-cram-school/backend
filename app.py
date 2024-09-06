from googleapiclient.errors import HttpError
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
import json
import os
from dotenv import load_dotenv
import service.transform_data
from service.spreadsheet_service import SpreadsheetService, DriveService
from fastapi import Request



# .envファイルから環境変数を読み込む
load_dotenv()

mapping_file = os.getenv('MAPPING_FILE')
mapping = service.transform_data.load_json(mapping_file)

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
def get_search_info():
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
def get_subjects(sheet_id: str):
    sp = SpreadsheetService(fileID=sheet_id)
    subject = sp.get_worksheet()
    return subject

# 一旦無効化
def user_info(sheet_id: str, subjects_id: int):
    sp = SpreadsheetService(fileID=sheet_id)
    # 指定されたシートIDに対応するシートの情報を取得
    subjects = sp.get_worksheet()
    # subjects_idを使用して、対応するシートの名前を検索
    for subject in subjects:
        if subject['value'] == subjects_id:
            sheet_name = subject['label']
            break
    else:
        return {"error": "Subject not found"}
    date_info = sp.closetDataFinder(worksheetname=sheet_name, start_row=2)
    # date_infoを大きい順に並び替える
    sorted_date_info = sorted(date_info, key=lambda x: x['row'], reverse=True)

    positions = [item['row'] for item in sorted_date_info]
    old_sheet = sp.get_old_sheet(postionCell=positions[0], worksheetname=sheet_name)
    mapping = service.transform_data.load_json(mapping_file)

    transformed_data = service.transform_data.transform_data(old_sheet[0], mapping)
    return json.dumps(transformed_data, ensure_ascii=False, indent=2)

@app.get('/search/subjects/reports/{sheet_id}/{subjects_id}')
def user_info_exp(sheet_id: str, subjects_id: str):
    try:
        sp = SpreadsheetService(fileID=sheet_id)
        subjects = sp.get_worksheet()
        sheet_name = None
        for subject in subjects:
            if subject['value'] == int(subjects_id):
                sheet_name = subject['label']
                break

        if sheet_name is None:
            return {"error": "Subject not found "}
        # シート内で取得する位置を探索
        date_info = sp.exp_DataFinder(sheetname=sheet_name, expotent_base=7, start_row=2)
        if date_info[0]['position'] != 0:
            # 過去の報告書が存在する場合の処理
            meticulous_date_info = sp.exp_DataFinder(sheetname=sheet_name, expotent_base=7, start_row=date_info[-1]['position'])
            liner_search = sp.closetDataFinder(sheetname=sheet_name, start_row=meticulous_date_info[-1]['position'])

            sorted_date_info = sorted(liner_search, key=lambda x: x['row'], reverse=True)
            positions = [item['row'] for item in sorted_date_info]

            # データ取得とデータの整形
            old_sheet = sp.get_old_sheet(postionCell=positions[0], sub_name=sheet_name)
            transformed_data = service.transform_data.transform_data(old_sheet[0], mapping)
            return transformed_data  # JSON形式としてそのまま返す
        else: # 過去の報告書が存在しない場合
            return service.transform_data.initialize_mapping_with_defaults(mapping)
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        return {"error": f"エラーが発生しました: {str(e)}"}


@app.post('/submit/report/{sheet_id}/{subjects_id}')
async def submit_report(sheet_id: str, subjects_id: str, request: Request):
    try:
        # JSON データを取得
        report_data = await request.json()  # awaitを使って結果を待つ
        # SpreadsheetService のインスタンスを作成
        sp = SpreadsheetService(fileID=sheet_id)
        subjects = sp.get_worksheet()
        # subjects_idを使用して、対応するシートの名前を検索
        for subject in subjects:
            if subject['value'] == int(subjects_id):
                sheet_name = subject['label']
                break
        else:
            return {"error": "Subject not found"}
        
        # シート内で入力する位置を探索
        date_info = sp.exp_DataFinder(sheetname=sheet_name, expotent_base=7, start_row=2)
        if date_info[0]['position'] != 0:
            # 過去の報告書が存在する場合
            meticulous_date_info = sp.exp_DataFinder(sheetname=sheet_name, expotent_base=7, start_row=date_info[-1]['position'])
            liner_search = sp.closetDataFinder(sheetname=sheet_name, start_row=meticulous_date_info[-1]['position'])

            # 最も左側の空セルの位置を決定
            target_position = liner_search[-1]['row'] + 6
        else:
            # 過去の報告書が存在しない場合
            # 最初の報告書に入力する最初の位置は2列目固定
            target_position = 1
        
        # JSONデータをスプレッドシート形式に変換
        mapping = service.transform_data.load_json(mapping_file)
        transformed_data = service.transform_data.reverse_transform_data(report_data, mapping)
        # データをスプレッドシートに登録
        sp.update_report(target_position, transformed_data, sheet_name=sheet_name)
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
        subjects = sp.get_worksheet()
        # subjects_idを使用して、対応するシートの名前を検索
        # subjects_idを使用して、対応するシートの名前を検索
        for subject in subjects:
            if subject['value'] == int(subjects_id):
                sheet_name = subject['label']
                break
        else:
            return {"error": "Subject not found"}
        
        # シート内で入力する位置を探索
        date_info = sp.exp_DataFinder(sheetname=sheet_name, expotent_base=7, start_row=2)
        meticulous_date_info = sp.exp_DataFinder(sheetname=sheet_name, expotent_base=7, start_row=date_info[-1]['position'])
        liner_search = sp.closetDataFinder(sheetname=sheet_name, start_row=meticulous_date_info[-1]['position'])

        # 最新の過去のデータの位置を取得
        target_position = liner_search[-1]['row']
        # JSONデータをスプレッドシート形式に変換
        mapping = service.transform_data.load_json(mapping_file)
        transformed_data = service.transform_data.reverse_transform_data(report_data, mapping)
        # データをスプレッドシートに登録
        sp.update_report(target_position, transformed_data, sheet_name=sheet_name)
        return {"status": "success", "message": "Report submitted successfully."}
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        return {"error": f"エラーが発生しました: {str(e)}"}

if __name__ == '__main__':
    #print(get_search_info())
    #x = get_subjects(sheet_id='1CEn2feUeQMfq885PVtyX96-ImOQxeqypeyePHpnPMg4')
    #print(x)
    #z = user_info_exp(sheet_id='1ShyFHM6cn9LB_QLpwQsToX3ipFkmc6ELrbXledRr7ic', subjects_id = 374409231)
    #print(z)
    
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)