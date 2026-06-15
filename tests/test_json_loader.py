import json

from app.infrastructure.json_loader import JSONLoader

def test_load_returns_query_result(tmp_path):
    """ load()メソッドがQueryResultインスタンスを返す"""
    # テスト用JSON
    data = {
        "columns": ["ID", "NAME"], 
        "rows": [
            [1, "Steve"], 
            [2, "Bruce"]
        ]
    }
    # 一時ファイル生成
    json_file = tmp_path / "test.json"
    # JSON保存
    with open(json_file, "w", encoding="utf-8") as f:
        json.dump(data, f)
    # テスト対象実行
    loader = JSONLoader()
    result = loader.load(json_file)

    # 検証
    assert result.columns == ["ID", "NAME"]
    assert result.rows == [
        [1, "Steve"], 
        [2, "Bruce"]
    ]