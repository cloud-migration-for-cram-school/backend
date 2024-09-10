import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from dotenv import load_dotenv
from service.transform_data import reverse_transform_data, load_json
from service.spreadsheet_service import SpreadsheetService, DriveService
from fastapi import Request
from fastapi.exceptions import HTTPException

# .envファイルから環境変数を読み込む
load_dotenv()

mapping_file = os.getenv('MAPPING_FILE')

def submit_report(sheet_id: str, subjects_id: int):
    try:
        # JSON データを取得
        report_data = {
            "basicInfo": {
                "dateAndTime": "8月24日(土)",
                "subjectName": "a",
                "teacherName": "a",
                "progressInSchool": "a",
                "homeworkProgress": "a",
                "homeworkAccuracy": "a"
            },
            "communication": {
                "forNextTeacher": "a",
                "fromDirector": "a"
            },
            "testReview": {
                "testAccuracy": "a",
                "classOverallStatus": "a",
                "rationale": ""
            },
            "lessonDetails": {
                "lessons": [
                    {
                        "material": "a",
                        "chapter": "a",
                        "accuracy": "100"
                    },
                    {
                        "material": "a",
                        "chapter": "a",
                        "accuracy": "100"
                    },
                    {
                        "material": "a",
                        "chapter": "a",
                        "accuracy": "100"
                    },
                    {
                        "material": "a",
                        "chapter": "a",
                        "accuracy": "100"
                    },
                    {
                        "material": "a",
                        "chapter": "a",
                        "accuracy": "100"
                    },
                    {
                        "material": "a",
                        "chapter": "a",
                        "accuracy": "100"
                    }
                ],
                "strengthsAndAreasForImprovement": "a"
            },
            "homework": {
                "assignments": [
                    {
                        "day": "8月25日(日)",
                        "tasks": [
                            {
                                "material": "a",
                                "rangeAndPages": "a"
                            },
                            {
                                "material": "a",
                                "rangeAndPages": "a"
                            }
                        ]
                    },
                    {
                        "day": "8月26日(月)",
                        "tasks": [
                            {
                                "material": "a",
                                "rangeAndPages": "a"
                            },
                            {
                                "material": "a",
                                "rangeAndPages": "a"
                            }
                        ]
                    },
                    {
                        "day": "8月27日(火)",
                        "tasks": [
                            {
                                "material": "a",
                                "rangeAndPages": "a"
                            },
                            {
                                "material": "a",
                                "rangeAndPages": "a"
                            }
                        ]
                    },
                    {
                        "day": "8月28日(水)",
                        "tasks": [
                            {
                                "material": "a",
                                "rangeAndPages": "a"
                            },
                            {
                                "material": "a",
                                "rangeAndPages": "a"
                            }
                        ]
                    },
                    {
                        "day": "8月29日(木)",
                        "tasks": [
                            {
                                "material": "a",
                                "rangeAndPages": "a"
                            },
                            {
                                "material": "a",
                                "rangeAndPages": "a"
                            }
                        ]
                    },
                    {
                        "day": "8月30日(金)",
                        "tasks": [
                            {
                                "material": "a",
                                "rangeAndPages": "a"
                            },
                            {
                                "material": "a",
                                "rangeAndPages": "a"
                            }
                        ]
                    }
                ],
                "advice": "a",
                "noteForNextSession": "a"
            },
            "nextTest": [
                {
                    "material": "CT",
                    "chapter": "a",
                    "rangeAndPages": "a"
                },
                {
                    "material": "CT",
                    "chapter": "a",
                    "rangeAndPages": "a"
                },
                {
                    "material": "CT",
                    "chapter": "a",
                    "rangeAndPages": "a"
                }
            ],
            "studentStatus": "a",
            "lessonPlan": {
                "ifTestOK": "a",
                "ifTestNG": "a"
            }
        }

        # SpreadsheetService のインスタンスを作成
        sp = SpreadsheetService(fileID=sheet_id)

        subjects = sp.get_worksheet()
        # subjects_idを使用して、対応するシートの名前を検索
        for subject in subjects:
            if subject['value'] == int(subjects_id):
                sheet_name = subject['label']
                break
        else:
            return {"error": "Subject not found"}

        # シート内で入力する位置を探索
        date_info = sp.exp_DataFinder(sheetname=sheet_name, expotent_base=7, start_row=2)
        meticulous_date_info = sp.exp_DataFinder(sheetname=sheet_name, expotent_base=7, start_row=date_info[-1]['position'])
        liner_search = sp.closetDataFinder(sheetname=sheet_name, start_row=meticulous_date_info[-1]['position'])

        # 最も左側の空セルの位置を決定
        target_position = liner_search[-1]['row'] + 6

        # JSONデータをスプレッドシート形式に変換
        mapping = load_json(mapping_file)
        transformed_data = reverse_transform_data(report_data, mapping)

        # データをスプレッドシートに登録
        sp.update_report(target_position, transformed_data, subject_id=sheet_name)

        return {"status": "success", "message": "Report submitted successfully."}
    except Exception as e:
        print(f"{e}")

if __name__ == "__main__":
    submit_report(sheet_id="1ShyFHM6cn9LB_QLpwQsToX3ipFkmc6ELrbXledRr7ic", subjects_id=374409231)