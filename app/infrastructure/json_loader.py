import json
from pathlib import Path

from app.models.query_result import QueryResult

class JSONLoader:
    def load(self, path: Path) -> QueryResult:
        with open(path, encoding="utf-8-sig") as fp:
            data = json.load(fp)
        return QueryResult(
            columns=data["columns"], 
            rows=data["rows"]
        )
