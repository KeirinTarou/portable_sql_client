import json
from pathlib import Path

from app.models.query_result import QueryResult

class JSONLoader:
    def load(self, path: Path) -> QueryResult:
        with open(path, encoding="utf-8-sig") as fp:
            data = json.load(fp)
        # DBエラーのとき
        #   踏み台Excelが`"column": "error"`のJSONを返す
        if data["columns"] == ["error"]:
            return QueryResult.error(
                title="DBからレコードセットが返らなかった", 
                message=data["rows"][0][0]
            )
        
        return QueryResult(
            columns=data["columns"], 
            rows=data["rows"]
        )
