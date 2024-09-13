from spreadsheet_service import SpreadsheetService
from transform_data import load_json
from dotenv import load_dotenv
import os
from googleapiclient.discovery import build

load_dotenv()
TEMPLATE_ID = os.getenv('TEMPLATE_ID')
newsheet_setting_path = os.getenv('NEWSHEET_SETTING')

NUMBER_OF_SHEET = 10 # 新しく作るシートの枚数

class MakeNewReport(SpreadsheetService):
    def __init__(self, fileID=None, position=None):
        super().__init__(fileID)
        self.fileID = fileID
        self.service = build('sheets', 'v4', credentials=self.credentials)
        self.newsheet_setting = load_json(newsheet_setting_path)
        self.position = position
    
    def copy_template(self):
        """
        テンプレートをコピーして貼り付ける処理をする
        """
        pass

    def apply_json_to_sheet(self):
        """
        Jsonファイルを読み込み、セル結合、セルの塗りつぶしを反映する
        """
        pass
    
    def debug_test(self):
        # batchUpdateリクエストの送信
        self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.fileID,
            body=self.newsheet_setting
        ).execute()
        print("Success")

if __name__ == '__main__':
    URL = '1NjVsRdbP1sQPgllo9wGYfpbmKjSWQ3us4_dMljvWN14'
    test = MakeNewReport(fileID=URL)
    test.debug_test()

"""
ガチの自分用のメモ
1,残りの枚数が少ないことをspreadsheet_service.py/find_nearby_datesでキャッチする
2, そのときの位置をMakeNewReportで受け取る
3, 位置を用いてテンプレートからコピーする
4, コピーしたらセル結合、フォントサイズ・色、セルの塗りつぶし、枠線の変更をする
5, NUMBER_OF_SHEETの枚数だけ作るようにする
"""