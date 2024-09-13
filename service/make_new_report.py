from service.spreadsheet_service import SpreadsheetService
from dotenv import load_dotenv
from service.transform_data import load_json
import os
from googleapiclient.discovery import build

load_dotenv()
TEMPLATE_ID = os.getenv('TEMPLATE_ID')

NUMBER_OF_SHEET = 10 # 新しく作るシートの枚数

class MakeNewReport(SpreadsheetService):
    def __init__(self, fileID=None, position=None):
        super().__init__(fileID)
        self.service = build('sheets', 'v4', credentials=self.credentials)
        self.position = position
    
    

"""
ガチの自分用のメモ
1,残りの枚数が少ないことをspreadsheet_service.py/find_nearby_datesでキャッチする
2, そのときの位置をMakeNewReportで受け取る
3, 位置を用いてテンプレートからコピーする
4, コピーしたらセル結合、フォントサイズ・色、セルの塗りつぶし、枠線の変更をする
5, NUMBER_OF_SHEETの枚数だけ作るようにする
"""