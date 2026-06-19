from typing import Tuple

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QPlainTextEdit, QTextEdit
from PyQt6.QtGui import (
    QTextCursor, QTextBlock, 
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
        """ カーソル行 or 選択行にインデントを追加する

        .. note::
        - カーソル点滅時
            - カーソル位置にTab（`\t`）を挿入
        - 選択時
            - 選択中の各行の先頭にTab（`\t`）を追加
            - 選択位置の最後が行頭のとき、最終行にはTabを追加しない
        """
        cursor = self.textCursor()
        # カーソルが選択状態
        if cursor.hasSelection():
            start_line, end_line = self._selected_line_range(cursor)
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
        """ カーソル行 or 選択行をアウトデントする

        .. note::
        - カーソル点滅時
            - カーソル位置のある行の先頭のTab（\t）を削除
            - 行頭が半角スペース4個の場合は半角スペース4個を削除
        - 選択時
            - 選択中の各行の先頭のTab（`\t`）または半角スペース4個を削除
            - 選択位置の最後が行頭のとき、最終行の行頭Tabや半角スペース4個は削除しない
        """
        cursor = self.textCursor()
        doc = cursor.document()
        if cursor.hasSelection():
            start_line, end_line = self._selected_line_range(cursor)
            for line_no in range(start_line, end_line + 1):
                block = doc.findBlockByNumber(line_no)
                self._remove_indent(cursor, block)    
        else:
            block = cursor.block()
            self._remove_indent(cursor, block)
    
    # keyPressイベントをフック
    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Tab:
            self._indent()
            return
        
        if event.key() == Qt.Key.Key_Backtab:
            self._outdent()
            return

        super().keyPressEvent(event)

    def _remove_indent(self, cursor: QTextCursor, block: QTextBlock):
        """ アウトデント用ヘルパ
        
        .. note::
        - ブロックの先頭がTabかスペース4個かに応じて適切に削除する
        """
        text = block.text()
        cursor.setPosition(block.position())
        del_len = 0
        if text.startswith("\t"):
            del_len = 1
        elif text.startswith("    "):
            del_len = 4
        cursor.movePosition(
            QTextCursor.MoveOperation.Right, 
            QTextCursor.MoveMode.KeepAnchor, 
            del_len
        )
        cursor.removeSelectedText()

    def _selected_line_range(
            self, 
            cursor: QTextCursor) -> Tuple[int, int]:
        """ 選択位置の開始行・終了行位置（0始まり）を返す
        
        .. note::
        - カーソル終端が行頭にあるときは、カーソル行は含めない
        """
        # エディタの文字列内におけるカーソルの相対位置取得
        start = cursor.selectionStart()
        end = cursor.selectionEnd()
        # カーソルの開始位置・終了位置が属する行位置（0始まり）を取得
        doc = cursor.document()
        start_line = (
            doc.findBlock(start).blockNumber()
        )
        end_line = (
            doc.findBlock(max(start, end - 1)).blockNumber()
        )
        return (start_line, end_line) 
