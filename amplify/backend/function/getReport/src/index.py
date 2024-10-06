import json
from service.spreadsheet_service import SpreadsheetService
from service.transform_data import transform_data 
from googleapiclient.errors import HttpError

def lambda_handler(event, context):
    if event['httpMethod'] == 'GET':
        try:
            sheetId = event['pathParameters']['sheetId']
            subjectId = event['pathParameters']['subjectId']
            
            sp = SpreadsheetService(fileID=sheetId)
            date_info = sp.find_exponential_dates(subject_id=int(subjectId), expotent_base=7, start_row=2)
            
            if date_info and date_info[0]['position'] != 0:
                meticulous_date_info = sp.find_exponential_dates(subject_id=int(subjectId), expotent_base=7, start_row=date_info[-1]['position'])
                positions, remaining_report_count = sp.find_closest_dates(subject_id=int(subjectId), start_row=meticulous_date_info[-1]['position'])

                if not positions:
                    return {
                        'statusCode': 500,
                        'body': json.dumps({"error": "positionsが空です。"})
                    }

                row_position = positions[0]['row']
                old_sheet = sp.get_old_sheet_data(postionCell=row_position, subject_id=int(subjectId))

                transformed_data = transform_data(old_sheet[0], remaining_report_count)
                
                return {
                    'statusCode': 200,
                    'headers': {
                        'Access-Control-Allow-Headers': '*',
                        'Access-Control-Allow-Origin': '*',
                        'Access-Control-Allow-Methods': 'OPTIONS,POST,GET'
                    },
                    'body': json.dumps(transformed_data)
                }
            else:
                default_mapping = transform_data.initialize_mapping_with_defaults()
                return {
                    'statusCode': 200,
                    'body': json.dumps(default_mapping)
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
                'body': json.dumps({'error': f"エラーが発生しました: {str(e)}"})
            }

    return {
        'statusCode': 400,
        'body': json.dumps({'error': 'Invalid request'})
    }
