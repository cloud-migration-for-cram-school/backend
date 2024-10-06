import json
from service.drive_service import DriveService
from googleapiclient.errors import HttpError

def lambda_handler(event, context):
    if event['httpMethod'] == 'GET':
        deiveservice = DriveService()
        try:
            items = deiveservice.get_info()
            file_info = [{'label': item['name'], 'value': item['id']} for item in items]
            
            return {
                'statusCode': 200,
                'headers': {
                    'Access-Control-Allow-Headers': '*',
                    'Access-Control-Allow-Origin': '*',
                    'Access-Control-Allow-Methods': 'OPTIONS,POST,GET'
                },
                'body': json.dumps(file_info)
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
