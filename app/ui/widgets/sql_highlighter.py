from PyQt6.QtGui import (
    QSyntaxHighlighter, 
    QTextCharFormat
)

class SQLHighlighter(QSyntaxHighlighter):
    def __init__(self, document):
        super().__init__(document)

    def highlightBlock(self, text):
        pass