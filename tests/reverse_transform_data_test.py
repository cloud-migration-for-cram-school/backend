import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from service.transform_data import reverse_transform_data, load_json
from dotenv import load_dotenv

# .envファイルから環境変数を読み込む
load_dotenv()

mapping_file = os.getenv('MAPPING_FILE')

# マッピングファイルの読み込み
mapping = load_json(mapping_file)

#疑似データの読み込み
sheet_data = {
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

print(reverse_transform_data(sheet_data, mapping))