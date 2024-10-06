import json
from service.spreadsheet_service import SpreadsheetService
from service.transform_data import reverse_transform_data
from googleapiclient.errors import HttpError

def lambda_handler(event, context):
    try:
        body = json.loads(event['body']) 
        
        sheetId = event['pathParameters']['sheetId']
        subjectId = event['pathParameters']['subjectId']
        
        sp = SpreadsheetService(fileID=sheetId)
        
        date_info = sp.find_exponential_dates(subject_id=int(subjectId), expotent_base=7, start_row=2)
        meticulous_date_info = sp.find_exponential_dates(subject_id=int(subjectId), expotent_base=7, start_row=date_info[-1]['position'])
        liner_search, remaining_report_count = sp.find_closest_dates(subject_id=int(subjectId), start_row=meticulous_date_info[-1]['position'])

        target_position = liner_search[0]['row'] - 1

        transformed_data = reverse_transform_data(body)
        sp.update_report(target_position, transformed_data, subject_id=int(subjectId))
        
        return {
            'statusCode': 200,
            'body': json.dumps({"status": "success", "message": "Report submitted successfully."})
        }
    
    except HttpError as error:
        message = f'エラー: {error.content.decode("utf-8")}'
        return {
            'statusCode': 500,
            'body': json.dumps({'error': message})
        }

    except Exception as e:
        return {
            'statusCode': 500,
            'body': json.dumps({"error": f"エラーが発生しました: {str(e)}"})
        }
