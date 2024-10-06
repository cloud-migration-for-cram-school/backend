import json
from service.spreadsheet_service import SpreadsheetService
from googleapiclient.errors import HttpError

def lambda_handler(event, context):
    if event['httpMethod'] == 'GET':
        try:
            sheetId = event['pathParameters']['sheetId']

            sp = SpreadsheetService(fileID=sheetId)
            subject = sp.get_worksheets()
            
            return {
                'statusCode': 200,
                'headers': {
                    'Access-Control-Allow-Headers': '*',
                    'Access-Control-Allow-Origin': '*',
                    'Access-Control-Allow-Methods': 'OPTIONS,POST,GET'
                },
                'body': json.dumps(subject)
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
