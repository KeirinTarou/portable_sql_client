from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QPlainTextEdit, QTextEdit
from PyQt6.QtGui import (
    QTextCursor, 
    QFont, QFontMetrics, 
    QColor, QTextFormat)

class SQLEditor(QPlainTextEdit):
    def __init__(self):
        super().__init__()
        # フォントの設定
        font = QFont("Consolas", 10)
        self.setFont(font)
        # タブ幅の調整
        metrics = QFontMetrics(font)
        self.setTabStopDistance(
            metrics.horizontalAdvance(" ") * 4
        )
        # 現在行ハイライト
        self.cursorPositionChanged.connect(
            self._highlight_current_line
        )
        # 現在行ハイライトを呼ぶ
        self._highlight_current_line()

    def _highlight_current_line(self):
        extra_selections = []

        if not self.isReadOnly():
            selection = QTextEdit.ExtraSelection()

            selection.format.setBackground(
                QColor(245, 245, 220)
            )

            selection.format.setProperty(
                QTextFormat.Property.FullWidthSelection, 
                True
            )

            selection.cursor = self.textCursor()
            selection.cursor.clearSelection()

            extra_selections.append(selection)
        
        self.setExtraSelections(
            extra_selections
        )

    def _indent(self):
        cursor = self.textCursor()
        # カーソルが選択状態
        if cursor.hasSelection():
            # エディタの文字列内におけるカーソルの相対位置取得
            start = cursor.selectionStart()
            end = cursor.selectionEnd()
            # カーソルの開始位置・終了位置が属する行位置（0始まり）を取得
            # QTextDocumentオブジェクト取得
            #   - エディタ内の全文を表すオブジェクト
            doc = cursor.document()
            start_line = (
                doc.findBlock(start).blockNumber()
            )
            end_line = (
                # カーソルの終端が行頭 -> 前行末に補正
                #   - 文字が選択されていない行を含めないため
                doc.findBlock(max(start, end - 1)).blockNumber()
            )
            for line_no in range(start_line, end_line + 1):
                block = cursor.document().findBlockByNumber(line_no)
                # 行の先頭位置にカーソルをセット
                cursor.setPosition(
                    # 行の先頭位置を取得
                    block.position()
                )
                # カーソル位置にタブを追加
                cursor.insertText("\t")
        # カーソルが点滅状態
        else:
            cursor.insertText("\t")

    def _outdent(self):
        cursor = self.textCursor()

        block_text = cursor.block().text()

        if not block_text.startswith("\t"):
            return

        cursor.movePosition(
            QTextCursor.MoveOperation.StartOfLine
        )

        cursor.deleteChar()
    
    # keyPressイベントをフック
    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Tab:
            self._indent()
            return
        
        if event.key() == Qt.Key.Key_Backtab:
            self._outdent()
            return

        super().keyPressEvent(event)
