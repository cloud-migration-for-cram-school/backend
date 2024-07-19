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