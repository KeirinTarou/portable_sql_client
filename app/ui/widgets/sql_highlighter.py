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
        keyword_rules = [
            ("SELECT", COLOR_CRIMSON), 
            ("FROM", COLOR_CRIMSON), 
            ("JOIN", COLOR_CRIMSON), 
            ("INNER", COLOR_CRIMSON), 
            ("OUTER", COLOR_CRIMSON), 
            ("CROSS", COLOR_CRIMSON), 
            ("LEFT", COLOR_CRIMSON), 
            ("RIGHT", COLOR_CRIMSON), 
            ("WHERE", COLOR_CRIMSON), 
            ("GROUP", COLOR_CRIMSON), 
            ("BY", COLOR_CRIMSON), 
            ("HAVING", COLOR_CRIMSON), 
            ("ORDER", COLOR_CRIMSON), 
            ("ASC", COLOR_CRIMSON), 
            ("DESC", COLOR_CRIMSON), 
            ("AS", COLOR_CRIMSON), 
            ("ON", COLOR_CRIMSON), 
            ("IN", COLOR_CRIMSON), 
        ]
        for rule in keyword_rules:
            keyword, color = rule
            self._add_keyword_rule(
                keyword, color
            )

    def highlightBlock(self, text):
        for pattern, fmt in self.rules:
            for match in pattern.finditer(text):
                self.setFormat(
                    match.start(), 
                    match.end() - match.start(), 
                    fmt
                )

    def _add_keyword_rule(self, keyword: str, color: str):
        fmt = QTextCharFormat()
        fmt.setForeground(
            QColor(color)
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