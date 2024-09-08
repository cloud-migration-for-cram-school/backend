import json
import os
from dotenv import load_dotenv
from googleapiclient.errors import HttpError
from service.spreadsheet_service import SpreadsheetService, DriveService
import service.transform_data

# 環境変数の読み込み
load_dotenv()
mapping_file = os.getenv('MAPPING_FILE')

# Lambdaのエントリポイント関数
def lambda_handler(event, context):
    """
    Lambdaハンドラ関数。API Gatewayからのリクエストを処理し、適切なレスポンスを返す。
    """
    path = event['path']
    method = event['httpMethod']

    # エンドポイントとメソッドに応じた処理の呼び出し
    if path == '/search' and method == 'GET':
        return get_search_info()

    elif path.startswith('/search/subjects/') and method == 'GET':
        sheet_id = path.split('/')[-1]
        return get_subjects(sheet_id)

    elif path.startswith('/search/subjects/reports/') and method == 'GET':
        path_parts = path.split('/')
        sheet_id = path_parts[-2]
        subjects_id = path_parts[-1]
        return user_info_exp(sheet_id, subjects_id)

    elif path.startswith('/submit/report/') and method == 'POST':
        path_parts = path.split('/')
        sheet_id = path_parts[-2]
        subjects_id = path_parts[-1]
        body = json.loads(event['body'])
        return submit_report(sheet_id, subjects_id, body)

    else:
        return {
            'statusCode': 404,
            'body': json.dumps({'message': 'Not Found'})
        }

# 各エンドポイントの処理

def get_search_info():
    """
    Google Driveのファイル情報を取得してレスポンスを返す。
    """
    try:
        drive_service = DriveService()
        items = drive_service.get_info()
        file_info = [{'label': item['name'], 'value': item['id']} for item in items]
        return {
            'statusCode': 200,
            'body': json.dumps(file_info)
        }
    except HttpError as error:
        return {
            'statusCode': 500,
            'body': json.dumps({'error': error.content.decode('utf-8')})
        }

def get_subjects(sheet_id: str):
    """
    指定されたスプレッドシートのIDに基づいて科目データを取得する。
    """
    try:
        sp = SpreadsheetService(fileID=sheet_id)
        subjects = sp.get_worksheet()
        return {
            'statusCode': 200,
            'body': json.dumps(subjects)
        }
    except Exception as e:
        return {
            'statusCode': 500,
            'body': json.dumps({'error': str(e)})
        }

def user_info_exp(sheet_id: str, subjects_id: str):
    """
    指定されたスプレッドシートIDと科目IDに基づいて、報告書データを取得し、変換する処理。
    """
    try:
        sp = SpreadsheetService(fileID=sheet_id)
        subjects = sp.get_worksheet()
        sheet_name = None
        for subject in subjects:
            if subject['value'] == int(subjects_id):
                sheet_name = subject['label']
                break

        if sheet_name is None:
            return {
                'statusCode': 404,
                'body': json.dumps({'error': 'Subject not found'})
            }

        date_info = sp.exp_DataFinder(sheetname=sheet_name, expotent_base=7, start_row=2)
        meticulous_date_info = sp.exp_DataFinder(sheetname=sheet_name, expotent_base=7, start_row=date_info[-1]['position'])
        liner_search = sp.closetDataFinder(sheetname=sheet_name, start_row=meticulous_date_info[-1]['position'])
        sorted_date_info = sorted(liner_search, key=lambda x: x['row'], reverse=True)
        positions = [item['row'] for item in sorted_date_info]
        old_sheet = sp.get_old_sheet(postionCell=positions[0], sub_name=sheet_name)
        mapping = service.transform_data.load_json(mapping_file)
        transformed_data = service.transform_data.transform_data(old_sheet[0], mapping)

        return {
            'statusCode': 200,
            'body': json.dumps(transformed_data)
        }
    except Exception as e:
        return {
            'statusCode': 500,
            'body': json.dumps({'error': str(e)})
        }

def submit_report(sheet_id: str, subjects_id: str, report_data):
    try:
        sp = SpreadsheetService(fileID=sheet_id)
        subjects = sp.get_worksheet()
        sheet_name = None

        for subject in subjects:
            if subject['value'] == int(subjects_id):
                sheet_name = subject['label']
                break
        if sheet_name is None:
            return {
                'statusCode': 404,
                'body': json.dumps({'error': 'Subject not found'})
            }
        # シート内で入力する位置を探索
        date_info = sp.exp_DataFinder(sheetname=sheet_name, expotent_base=7, start_row=2)
        meticulous_date_info = sp.exp_DataFinder(sheetname=sheet_name, expotent_base=7, start_row=date_info[-1]['position'])
        liner_search = sp.closetDataFinder(sheetname=sheet_name, start_row=meticulous_date_info[-1]['position'])
        # 最も左側の空セルの位置を決定
        target_position = liner_search[-1]['row'] + 6
        # JSONデータをスプレッドシート形式に変換
        mapping = service.transform_data.load_json(mapping_file)
        transformed_data = service.transform_data.reverse_transform_data(report_data, mapping)
        # データをスプレッドシートに登録
        sp.update_report(target_position, transformed_data, sheet_name=sheet_name)
        return {
            'statusCode': 200,
            'body': json.dumps({'status': 'success', 'message': 'Report submitted successfully.'})
        }

    except Exception as e:
        return {
            'statusCode': 500,
            'body': json.dumps({'error': str(e)})
        }
