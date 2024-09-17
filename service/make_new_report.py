from spreadsheet_service import SpreadsheetService
from transform_data import load_json
from dotenv import load_dotenv
import os
from googleapiclient.discovery import build
import json
import os

load_dotenv()
TEMPLATE_ID = os.getenv('TEMPLATE_ID')
newsheet_setting_path = os.getenv('NEWSHEET_SETTING')

NUMBER_OF_SHEET = 10 # 新しく作るシートの枚数

class MakeNewReport(SpreadsheetService):
    def __init__(self, fileID=None, sheetID=None, position=None):
        super().__init__(fileID)
        self.fileID = fileID
        self.service = build('sheets', 'v4', credentials=self.credentials)
        self.newsheet_setting = load_json(newsheet_setting_path)
        self.position = position
        self.sheetID = sheetID

    def apply_json_to_sheet(self):
        """
        Jsonファイルを読み込み、セル結合、セルの塗りつぶしを反映する
        """
        self.molding_json()  # JSONを整形
        # batchUpdateリクエストの送信
        self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.fileID,
            body=self.newsheet_setting
        ).execute()

    def molding_json(self):
        """
        Jsonファイルを整形する関数
        """
        for request in self.newsheet_setting["requests"]:
            try:
                # mergeCells の処理
                if "range" in request.get("mergeCells", {}):
                    request["mergeCells"]["range"]["sheetId"] = self.sheetID
                    start_col = request["mergeCells"]["range"]["startColumnIndex"]
                    end_col = request["mergeCells"]["range"]["endColumnIndex"]
                    # 負の値を避けるためにmax()を使ってインデックスを修正
                    request["mergeCells"]["range"]["startColumnIndex"] = max(0, self.position + start_col)
                    request["mergeCells"]["range"]["endColumnIndex"] = max(0, self.position + end_col)
                
                # repeatCell の処理
                elif "range" in request.get("repeatCell", {}):
                    request["repeatCell"]["range"]["sheetId"] = self.sheetID
                    start_col = request["repeatCell"]["range"]["startColumnIndex"]
                    end_col = request["repeatCell"]["range"]["endColumnIndex"]
                    # 負の値を避けるためにmax()を使ってインデックスを修正
                    request["repeatCell"]["range"]["startColumnIndex"] = max(0, self.position + start_col)
                    request["repeatCell"]["range"]["endColumnIndex"] = max(0, self.position + end_col)
                    
            except Exception as e:
                print(e)
                continue


if __name__ == '__main__':
    URL = '1NjVsRdbP1sQPgllo9wGYfpbmKjSWQ3us4_dMljvWN14'
    test = MakeNewReport(fileID=URL, sheetID=1300349524, position = 8)
    test.apply_json_to_sheet()  # 整形後にAPIに送信


"""
ガチの自分用のメモ
1,残りの枚数が少ないことをspreadsheet_service.py/find_nearby_datesでキャッチする
2, そのときの位置をMakeNewReportで受け取る
3, 位置を用いてテンプレートからコピーする
4, コピーしたらセル結合、フォントサイズ・色、セルの塗りつぶし、枠線の変更をする
5, NUMBER_OF_SHEETの枚数だけ作るようにする
"""