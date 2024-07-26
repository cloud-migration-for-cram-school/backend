from googleapiclient.errors import HttpError
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
import json
import os
from dotenv import load_dotenv
import service.transform_data
from service.spreadsheet_service import SpreadsheetService
import time

# .envファイルから環境変数を読み込む
load_dotenv()

mapping_file = os.getenv('MAPPING_FILE')

#fastapiの有効化
app = FastAPI(debug=True)

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
    sp = SpreadsheetService()
    try:
        items = sp.get_info()
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


@app.get('/search/subjects/reports/{subjects_id}}')
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
    print(f'科目名 : {sheet_name}')
    date_info = sp.closetDataFinder(sheetname=sheet_name, start_row=2)
    # date_infoを大きい順に並び替える
    sorted_date_info = sorted(date_info, key=lambda x: x['row'], reverse=True)

    positions = [item['row'] for item in sorted_date_info]
    old_sheet = sp.get_old_sheet(postionCell=positions[0], sub_name=sheet_name)
    mapping = service.transform_data.load_json(mapping_file)

    transformed_data = service.transform_data.transform_data(old_sheet[0], mapping)
    return json.dumps(transformed_data, ensure_ascii=False, indent=2)


@app.get('/search/subjects/reports/{subjects_id}}')
def user_info_exp(sheet_id: str, subjects_id: int):
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
    print(f'科目名 : {sheet_name}')
    # epotent_baseは7で固定する
    # start_rowは2で固定する
    date_info = sp.exp_DataFinder(sheetname = sheet_name, expotent_base = 7, start_row=2)
    # 前回の指数探索で取得したデータからスタートをする
    meticulous_date_info = sp.exp_DataFinder(sheetname = sheet_name, expotent_base = 7, start_row=date_info[len(date_info) - 1]['position'])
    liner_search = sp.closetDataFinder(sheetname = sheet_name, start_row=meticulous_date_info[len(meticulous_date_info) - 1]['position'])

    # date_infoを大きい順に並び替える
    sorted_date_info = sorted(liner_search, key=lambda x: x['row'], reverse=True)
    positions = [item['row'] for item in sorted_date_info]
    old_sheet = sp.get_old_sheet(postionCell=positions[0], sub_name=sheet_name)
    mapping = service.transform_data.load_json(mapping_file)

    transformed_data = service.transform_data.transform_data(old_sheet[0], mapping)
    return json.dumps(transformed_data, ensure_ascii=False, indent=2)



if __name__ == '__main__':
    start_time = time.time()  # API呼び出し前の時刻を記録
    #print(get_search_info())
    #x = get_subjects(sheet_id='1CEn2feUeQMfq885PVtyX96-ImOQxeqypeyePHpnPMg4')
    #print(x)
    z = user_info_exp(sheet_id='1CEn2feUeQMfq885PVtyX96-ImOQxeqypeyePHpnPMg4', subjects_id = 1059696948)
    print(z)

    end_time = time.time()  # API呼び出し後の時刻を記録
    elapsed_time = end_time - start_time  # 応答時間を計算
    print(f"API応答時間: {elapsed_time}秒")
    # 変更前 5秒 変更後 4.3秒