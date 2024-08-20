import json

def load_json(file_path):
    """JSON ファイルを読み込む"""
    with open(file_path, 'r', encoding='utf-8') as file:
        return json.load(file)

def transform_data(input_data, mapping):
    """入力データをマッピングに基づいて変換する"""
    transformed_data = {}
    for key, value in mapping.items():
        if isinstance(value, list):
            # リスト内の要素が辞書の場合、再帰的に処理
            if all(isinstance(item, dict) for item in value):
                transformed_data[key] = [transform_data(input_data, item) for item in value]
            else:
                try:
                    row, col = value
                    # input_dataのサイズをチェック
                    if row < len(input_data) and col < len(input_data[row]):
                        transformed_data[key] = input_data[row][col]
                    else:
                        # 範囲外の場合、デフォルト値を設定
                        transformed_data[key] = None
                except ValueError:
                    # アンパックできない場合のエラー処理
                    print(f"Error processing key {key}: cannot unpack {value}")
        elif isinstance(value, dict):
            # ネストされた構造の処理
            transformed_data[key] = transform_data(input_data, value)
        else:
            # 予期せぬ形式の場合
            transformed_data[key] = None
    return transformed_data

def reverse_transform_data(input_data, mapping):
    """入力データをスプレッドシート形式に変換する"""
    reversed_data = [[None] * 6 for _ in range(32)]

    for key, value in mapping.items():
        if isinstance(value, list):
            if all(isinstance(item, dict) for item in value):
                for sub_item, lesson in zip(value, input_data[key]['lessons']):
                    for sub_key, sub_value in sub_item.items():
                        row, col = sub_value
                        reversed_data[row][col] = lesson[sub_key]
            else:
                row, col = value
                reversed_data[row][col] = input_data[key]
        elif isinstance(value, dict):
            if key == "assignments":
                for assignment, day_info in zip(value['assignments'], input_data['homework']['assignments']):
                    row, col = assignment['day']
                    reversed_data[row][col] = day_info['day']
                    for task, task_info in zip(assignment['tasks'], day_info['tasks']):
                        task_row, task_col = task['material']
                        reversed_data[task_row][task_col] = task_info['material']
                        range_row, range_col = task['rangeAndPages']
                        reversed_data[range_row][range_col] = task_info['rangeAndPages']
            else:
                reversed_data = reverse_transform_data(input_data[key], value)
        else:
            reversed_data[row][col] = input_data[key]

    return reversed_data

def main():
    # マッピングと入力データを読み込む
    mapping = load_json('mapping.json')
    input_data = load_json('input_data.json')  # 仮の入力データファイル名

    # データ変換
    transformed_data = transform_data(input_data, mapping)

    # 変換結果をファイルに出力
    with open('transformed_data.json', 'w', encoding='utf-8') as file:
        json.dump(transformed_data, file, ensure_ascii=False, indent=4)

if __name__ == '__main__':
    main()