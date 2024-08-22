import unittest
import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from service.transform_data import reverse_transform_data, transform_data, load_json
from dotenv import load_dotenv



class TestTransformData(unittest.TestCase):

    def setUp(self):

        # .envファイルから環境変数を読み込む
        load_dotenv()
        mapping_file = os.getenv('MAPPING_FILE')
        self.mapping = load_json(mapping_file)

        expected_output = {
            "basicInfo": {
                "dateAndTime": "2022-03-15",
                "subjectName": "数学",
                "teacherName": "Taro Yamada",
                "progressInSchool": "5th chapter",
                "homeworkProgress": 80,
                "homeworkAccuracy": 80
            },
            "communication": {
                "forNextTeacher": "insert message for teacher",
                "fromDirector": "insert message from director"
            },
            "testReview": {
                "testAccuracy": 60,
                "classOverallStatus": "good",
                "rationale": "insert reason"
            },
            "lessonDetails": {
                "lessons": [
                    {
                        "material": "Scramble",
                        "chapter": "5th chapter",
                        "accuracy": 40
                    },
                    {
                        "material": "Vintage",
                        "chapter": "7th chapter",
                        "accuracy": 80
                    },
                    {
                        "material": "教科書",
                        "chapter": "第7章 分数",
                        "accuracy": 90
                    },
                    {
                        "material": "教科書",
                        "chapter": "第8章 分数",
                        "accuracy": 90
                    },
                    {
                        "material": "教科書",
                        "chapter": "第9章 分数",
                        "accuracy": 90
                    },
                    {
                        "material": "教科書",
                        "chapter": "第10章 分数",
                        "accuracy": 90
                    }
                ],
                "strengthsAndAreasForImprovement": "reasoning was good"
            },
            "homework": {
                "assignments": [
                    {
                        "day": "2022-03-16",
                        "tasks": [
                            {
                                "material": "textbook",
                                "rangeAndPages": "p.100-105"
                            },
                            {
                                "material": "workbook",
                                "rangeAndPages": "p.20-25"
                            }
                        ]
                    },
                    {
                        "day": "2022-03-17",
                        "tasks": [
                            {
                                "material": "教科書",
                                "rangeAndPages": "p.106-110"
                            },
                            {
                                "material": "ワークブック",
                                "rangeAndPages": "p.26-30"
                            }
                        ]
                    },
                    {
                        "day": "2022-03-18",
                        "tasks": [
                            {
                                "material": "教科書",
                                "rangeAndPages": "p.111-115"
                            },
                            {
                                "material": "ワークブック",
                                "rangeAndPages": "p.31-35"
                            }
                        ]
                    },
                    {
                        "day": "2022-03-19",
                        "tasks": [
                            {
                                "material": "教科書",
                                "rangeAndPages": "p.116-120"
                            },
                            {
                                "material": "ワークブック",
                                "rangeAndPages": "p.36-40"
                            }
                        ]
                    },
                    {
                        "day": "2022-03-20",
                        "tasks": [
                            {
                                "material": "教科書",
                                "rangeAndPages": "p.121-125"
                            },
                            {
                                "material": "ワークブック",
                                "rangeAndPages": "p.41-45"
                            }
                        ]
                    },
                    {
                        "day": "2022-03-21",
                        "tasks": [
                            {
                                "material": "教科書",
                                "rangeAndPages": "p.126-130"
                            },
                            {
                                "material": "ワークブック",
                                "rangeAndPages": "p.46-50"
                            }
                        ]
                    }
                ],
                "advice": "refer to textbook",
                "noteForNextSession": "compass"
            },
            "nextTest": [
                {
                    "material": "CT",
                    "chapter": "6th chapter",
                    "rangeAndPages": "p.110-115"
                },
                {
                    "material": "LCT",
                    "chapter": "第6章 比率",
                    "rangeAndPages": "p.110-115"
                },
                {
                    "material": "PT",
                    "chapter": "第6章 比率",
                    "rangeAndPages": "p.110-115"
                }
            ],
            "studentStatus": "good",
            "lessonPlan": {
                "ifTestOK": "turn to the next chapter",
                "ifTestNG": "review the previous chapter"
            }
        }

        self.input_data = reverse_transform_data(expected_output, self.mapping)

    def test_transform_data(self):
        self.maxDiff = None
        
        transformed_data = transform_data(self.input_data, self.mapping)
      
        print(self.input_data)
        #self.assertEqual(transformed_data, expected_output)
        print(transformed_data)

if __name__ == '__main__':
    unittest.main()