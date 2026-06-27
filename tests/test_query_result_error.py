from app.models.query_result import QueryResult
from app.core.text_utils import truncate

def test_error_creates_error_query_result():
    """ error()メソッドでエラー表示用のQueryResultインスタンスが作れる"""
    title = "IRON_MAIDEN"
    message = "Fear_Of_The_Dark"
    res = QueryResult.error(title, message)
    assert res.columns == ["( ´,_ゝ｀)", "ち～ん（笑）"]
    assert res.rows == [
        ["残念ｗ", f"{title}ｗｗｗ"], 
        ["原因はたぶん……", truncate(message)], 
    ]