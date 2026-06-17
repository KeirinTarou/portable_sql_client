from dataclasses import dataclass
from typing import Any

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
    is_error: bool = False
    error_message: str = ""

    @property
    def row_count(self) -> int:
        return len(self.rows)
    
    @property
    def column_count(self) -> int:
        return len(self.columns)
    