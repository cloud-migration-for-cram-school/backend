import service.spreadsheet_service_init
from service.transform_data import load_json
from dotenv import load_dotenv
import os
from googleapiclient.discovery import build
import os

load_dotenv()
TEMPLATE_ID = os.getenv('TEMPLATE_ID')
newsheet_setting_path = os.getenv('NEWSHEET_SETTING')

NUMBER_OF_SHEET = 10 # 新しく作るシートの枚数
MARGE_COLUMN = 7 # シートの間隔

class MakeNewReport(service.spreadsheet_service_init.SpreadsheetServiceInit):
    def __init__(self, fileID=None, sheetID=None, position=None):
        super().__init__(fileID)
        self.fileID = fileID
        self.service = build('sheets', 'v4', credentials=self.credentials)
        self.position = position - 1  # 微調整
        self.sheetID = sheetID

    def apply_json_to_sheet(self):
        """
        Jsonファイルを読み込み、セル結合、セルの塗りつぶしを反映する
        """
        self.add_columns()
        for i in range(0, NUMBER_OF_SHEET):
            self.molding_json(i * MARGE_COLUMN)
            # batchUpdateリクエストの送信
            self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.fileID,
                body=self.newsheet_setting
            ).execute()

    def molding_json(self, iteration):
        """
        Jsonファイルを整形する関数
        iteration: 何回目の処理かに応じてシートIDやセルの位置を変える
        """
        self.newsheet_setting = load_json(newsheet_setting_path)
        for request in self.newsheet_setting["requests"]:
            try:
                # mergeCells の処理
                if "range" in request.get("mergeCells", {}):
                    request["mergeCells"]["range"]["sheetId"] = self.sheetID
                    start_col = request["mergeCells"]["range"]["startColumnIndex"]
                    end_col = request["mergeCells"]["range"]["endColumnIndex"]
                    # 負の値を避けるためにmax()を使ってインデックスを修正
                    request["mergeCells"]["range"]["startColumnIndex"] = self.position + start_col + iteration
                    request["mergeCells"]["range"]["endColumnIndex"] = self.position + end_col + iteration

                # repeatCell の処理
                elif "range" in request.get("repeatCell", {}):
                    request["repeatCell"]["range"]["sheetId"] = self.sheetID
                    start_col = request["repeatCell"]["range"]["startColumnIndex"]
                    end_col = request["repeatCell"]["range"]["endColumnIndex"]
                    # 負の値を避けるためにmax()を使ってインデックスを修正
                    request["repeatCell"]["range"]["startColumnIndex"] = self.position + start_col + iteration
                    request["repeatCell"]["range"]["endColumnIndex"] = self.position + end_col + iteration

                # updateBorders の処理
                elif "range" in request.get("updateBorders", {}):
                    request["updateBorders"]["range"]["sheetId"] = self.sheetID
                    start_col = request["updateBorders"]["range"]["startColumnIndex"]
                    end_col = request["updateBorders"]["range"]["endColumnIndex"]
                    # 負の値を避けるためにmax()を使ってインデックスを修正
                    request["updateBorders"]["range"]["startColumnIndex"] = self.position + start_col + iteration
                    request["updateBorders"]["range"]["endColumnIndex"] = self.position + end_col + iteration

            except Exception as e:
                print(f"Error in iteration {iteration}: {e}")
                continue

    def add_columns(self):
        """
        Add columns to the specified sheet.
        Adjusts the number of columns added based on the current grid size and position.
        """
        # Get the current grid size
        sheet_metadata = self.service.spreadsheets().get(spreadsheetId=self.fileID).execute()
        sheets = sheet_metadata.get('sheets', '')
        
        max_columns = 0

        for sheet in sheets:
            if sheet['properties']['sheetId'] == self.sheetID:
                grid_properties = sheet['properties']['gridProperties']
                max_columns = grid_properties['columnCount']  # Get the current number of columns
                print('max_columns:', max_columns)

        try:
            requests = [{
                "insertDimension": {
                    "range": {
                        "sheetId": self.sheetID,
                        "dimension": "COLUMNS",
                        "startIndex": max_columns - 1,
                        "endIndex": max_columns + NUMBER_OF_SHEET * MARGE_COLUMN - 1
                    },
                    "inheritFromBefore": False
                }
            }]

            body = {
                'requests': requests
            }

            # batchUpdateリクエストの送信
            self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.fileID,
                body=body
            ).execute()

            print(f"Added {NUMBER_OF_SHEET} columns to sheet with ID {self.sheetID}")
        
        except Exception as e:
            print(e)
