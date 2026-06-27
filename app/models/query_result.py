# 前方参照用
from __future__ import annotations
from dataclasses import dataclass
from typing import Any

from app.core.text_utils import truncate

"""
`@dataclass`デコレータにより、下記のようなコンストラクタが
裏で作成されている。
```
class QueryResult:
    def __init__(
        self,
        columns: list[str],
        rows: list[list[Any]]
    ):
        self.columns = columns
        self.rows = rows
```
"""

@dataclass
class QueryResult:
    columns: list[str]
    rows: list[list[Any]]

    @property
    def row_count(self) -> int:
        return len(self.rows)
    
    @property
    def column_count(self) -> int:
        return len(self.columns)
    
    @classmethod
    def error(cls, title: str, message: str) -> QueryResult:
        return cls(
            columns=["( ´,_ゝ｀)", "ち～ん（笑）"], 
            rows=[
                ["残念ｗ", f"{title}ｗｗｗ"], 
                ["原因はたぶん……", f"{truncate(message)}"]
            ], 
        )