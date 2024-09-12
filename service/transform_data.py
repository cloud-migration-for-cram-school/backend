import json
from dotenv import load_dotenv
import os

def load_json(file_path):
    """JSON ファイルを読み込む"""
    with open(file_path, 'r', encoding='utf-8') as file:
        return json.load(file)
    
load_dotenv()

mapping_file_path = os.getenv('MAPPING_FILE')
MAPPING = load_json(mapping_file_path)

def transform_data(input_data, reports_left, mapping=MAPPING):
    """入力データをマッピングに基づいて変換し、全てを文字列として処理する"""
    transformed_data = {}
    
    # spreadsheetInfo の部分を事前に追加しておく
    transformed_data["spreadsheetInfo"] = {
        "remainingReports": str(reports_left)
    }
    
    for key, value in mapping.items():
        if key == "spreadsheetInfo":
            # spreadsheetInfoの処理はスキップする
            continue
        
        if isinstance(value, list):
            # リスト内の要素が辞書の場合、再帰的に処理
            if all(isinstance(item, dict) for item in value):
                transformed_data[key] = [transform_data(input_data, reports_left, item) for item in value]
            else:
                try:
                    row, col = value
                    # input_dataのサイズをチェック
                    if row < len(input_data) and col < len(input_data[row]):
                        transformed_data[key] = str(input_data[row][col])
                    else:
                        # 範囲外の場合、デフォルト値を設定
                        transformed_data[key] = None
                except ValueError:
                    # アンパックできない場合のエラー処理
                    print(f"Error processing key {key}: cannot unpack {value}")
        elif isinstance(value, dict):
            # ネストされた構造の処理
            transformed_data[key] = transform_data(input_data, reports_left, value)
        else:
            # 予期せぬ形式の場合
            transformed_data[key] = None
    return transformed_data


def reverse_transform_data(input_data, mapping=MAPPING):
    """入力データをスプレッドシート形式に変換する"""
    reversed_data = [[None] * 6 for _ in range(37)] 

    for key, value in mapping.items():
        if isinstance(value, list):
            if all(isinstance(item, dict) for item in value):#"nextTest"の処理
                for item, test in zip(value, input_data[key]):
                    for sub_key, sub_value in item.items():
                        row, col = sub_value
                        reversed_data[row][col] = test[sub_key]
            else:#"studentStatus"の処理
                row, col = value
                reversed_data[row][col] = input_data[key]
        elif isinstance(value, dict):#"basicInfo"系と"lessonDetails", "homework"が通る
            for sub_key, sub_value in value.items():
                if sub_key == "lessons":
                    for item, lesson in zip(sub_value, input_data[key][sub_key]):
                        for sub_sub_key, sub_sub_value in item.items():
                            row, col = sub_sub_value
                            reversed_data[row][col] = lesson[sub_sub_key]
                elif sub_key == "assignments":
                    for item, assignment in zip(sub_value, input_data['homework']['assignments']):
                        row, col = item['day']
                        reversed_data[row][col] = assignment['day']
                        for task, task_info in zip(item['tasks'], assignment['tasks']):
                            task_row, task_col = task['material']
                            reversed_data[task_row][task_col] = task_info['material']
                            range_row, range_col = task['rangeAndPages']
                            reversed_data[range_row][range_col] = task_info['rangeAndPages']
                else:#"basicInfo"系の処理
                    row, col = sub_value
                    reversed_data[row][col] = input_data[key][sub_key]
    return reversed_data

def initialize_mapping_with_defaults(mapping=MAPPING):
    """
    入力データをマッピングに基づいて変換し、すべてを空の文字列で初期化し、特定のキーにデフォルト値を設定する
    """
    transformed_data = {}
    for key, value in mapping.items():
        if isinstance(value, list):
            transformed_data[key] = []
            for item in value:
                if isinstance(item, dict):
                    transformed_data[key].append(initialize_mapping_with_defaults(item))
                else:
                    transformed_data[key].append("")  # リストの要素を空の文字列に初期化
        elif isinstance(value, dict):
            transformed_data[key] = initialize_mapping_with_defaults(value)
        else:
            transformed_data[key] = ""  # その他の値を空の文字列に初期化

    # 特定のキーに対するデフォルト値を設定
    if 'communication' in transformed_data:
        transformed_data['communication']['forNextTeacher'] = "初回授業です"
        transformed_data['communication']['fromDirector'] = ""

    return transformed_data