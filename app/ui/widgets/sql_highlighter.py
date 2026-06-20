import re

from PyQt6.QtGui import (
    QSyntaxHighlighter, 
    QTextCharFormat, QColor 
)

COLOR_MAROON = "#800000"
COLOR_CRIMSON = "#dc143c"

class SQLHighlighter(QSyntaxHighlighter):
    def __init__(self, document):
        super().__init__(document)

        self.rules = []
        keywords = [
            "SELECT", "FROM", "JOIN", "INNER", "OUTER", 
            "CROSS", "LEFT", "RIGHT", "WHERE", "GROUP", 
            "BY", "HAVING", "ORDER", "ASC", "DESC", 
            "AS", "ON", "IN", "LIKE", "OR", 
            "AND", "NOT", "CASE", "WHEN", "THEN", 
            "ELSE", "END", "NULL"
        ]
        # キーワード・演算子用ルールを追加
        for kw in keywords:
            self._add_keyword_rule(kw)

    def highlightBlock(self, text):
        for pattern, fmt in self.rules:
            for match in pattern.finditer(text):
                self.setFormat(
                    match.start(), 
                    match.end() - match.start(), 
                    fmt
                )

    def _add_keyword_rule(self, keyword: str):
        fmt = QTextCharFormat()
        fmt.setForeground(
            QColor(COLOR_CRIMSON)
        )

        self.rules.append(
            (
                re.compile(
                    rf"\b{keyword}\b", 
                    re.IGNORECASE
                ), 
                fmt, 
            )
        )
