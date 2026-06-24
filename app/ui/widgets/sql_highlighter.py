import re
from enum import IntEnum
from typing import Tuple, List

from PyQt6.QtGui import (
    QSyntaxHighlighter, 
    QTextCharFormat, QColor 
)

COLOR_MAROON = "#800000"
COLOR_FORESTGREEN = "#228b22"
COLOR_ROYALBLUE = "#4169E1"
COLOR_CRIMSON = "#dc143c"
COLOR_SLATEBLUE = "#6A5ACD"
COLOR_CORAL = "#FF7F50"
COLOR_DEEPPINK = "#FF1493"
COLOR_DARKGRAY = "#a9a9a9"

class LexerState(IntEnum):
    NORMAL = 0
    STRING = 1
    BLOCK_COMMENT = 2

def _analyze_line(
            text: str, prev_state: LexerState) -> Tuple[LexerState, List[List[int]]]:
        """ 行（ブロック）の状態を判定する
        
        :param text: ブロックの文字列
        :type text: str
        :param prev_state: 直近のブロックの状態
        :type prev_state: LexerState
        :return: 次のブロックに引き継ぐ状態と、コメント領域の
        [start, end]リストのリストのタプル

        .. note::
        - app/ui/widgets/sql_highlighter.py
        """
        # 状態にかかわらず、クオートされた文字列は`_`に置換する
        text = _mask_quoted_parts(text)
        # 現在通常状態
        if prev_state == LexerState.NORMAL:
            start = text.find("/*")
            regions = []
            state = LexerState.NORMAL
            while start > -1:
                # `*/`を探す
                end = text.find("*/", start)
                # `*/`がない -> コメントが閉じられない
                #   -> 残り全部コメント・ステータス変更 -> exit
                if end == -1:
                    end = len(text)
                    state = LexerState.BLOCK_COMMENT
                    regions.append([start, end])
                    break
                # `*/`あり -> 位置情報追加 -> start更新
                else:
                    # `*/`の長さ分プラス
                    end += len("*/")
                    regions.append([start, end])
                    start = text.find("/*", end)
            return (state, regions)
        # 現在ブロックコメント内
        elif prev_state == LexerState.BLOCK_COMMENT:
            start = 0
            end = text.find("*/")
            regions = []
            state = LexerState.BLOCK_COMMENT
            # 閉じ`*/`が見つかっていたらループ突入
            while end != -1:
                end += len("*/")
                regions.append([start, end])
                # 次の`/*`を探す
                start = text.find("/*", end)
                # 次の`/*`あり
                if start != -1:
                    state = LexerState.BLOCK_COMMENT
                    # 閉じる`*/`を探して次へ
                    end = text.find("*/", start + len("/*"))
                # 次の`/*`なし -> ループを抜ける
                else:
                    state = LexerState.NORMAL
                    end = len(text)
                    break
            # 閉じ`*/`が見つからなかった -> 行末位置に補正
            if end == -1: 
                end = len(text)
                regions.append([start, end])
            return (state, regions)
        
def _mask_quoted_parts(text: str) -> str:
    """ クォートされた部分を引用ごと`_`に置き換える"""
    # シングルクォートの置換
    patt_single_quoted = r"'(?:''|[^'])*'"
    res = \
        re.sub(
            patt_single_quoted, 
            lambda m: "_" * len(m.group()), 
            text
        )
    patt_double_quoted = r'"(?:[^"]|"")*"'
    res = \
        re.sub(
            patt_double_quoted, 
            lambda m: "_" * len(m.group()), 
            res
        )
    
    return res

class SQLHighlighter(QSyntaxHighlighter):
    def __init__(self, document):
        super().__init__(document)

        self.rules = []
        # キーワード・演算子用ルールを追加
        keywords = [
            "SELECT", "FROM", "JOIN", "INNER", "OUTER", 
            "CROSS", "LEFT", "RIGHT", "WHERE", "GROUP", 
            "BY", "HAVING", "ORDER", "ASC", "DESC", 
            "AS", "ON", "IN", "LIKE", "OR", 
            "AND", "NOT", "CASE", "WHEN", "THEN", 
            "ELSE", "END", "IS", "BETWEEN", "DISTINCT",   
            "WITH", "UNION", "ALL", "EXISTS", 
            "WITHIN",  
        ]
        for kw in keywords:
            self._add_keyword_rule(kw, COLOR_CRIMSON)

        # 準キーワード用ルールを追加
        sub_keywords = [
            "NULL", "TRUE", "FALSE", "SYSDATE", "SYSTIMESTAMP"
        ]
        for kw in sub_keywords:
            self._add_keyword_rule(kw, COLOR_CORAL)

        # 関数用ルールを追加
        func_names = [
            "COUNT", "AVG", "SUM", "MAX", "MIN", 
            "TO_CHAR", "TO_DATE", "TO_TIMESTAMP" 
            "COALESCE", "NVL", "NVL2", "GREATEST", "LEAST", 
            "LISTAGG", "ADD_MONTHS", "MONTHS_BETWEEN", "TRUNC"
            "SUBSTR", "INSTR", "REPLACE", "TRIM", 
            "UPPER", "LOWER", "LPAD", "RPAD", 
            "ROUND", "CEIL", "FLOOR"
        ]
        for kw in func_names:
            self._add_function_rule(kw)

        # コメント用ルールを追加
        comment_patterns = [
            r"--.*$"
        ]
        for pattern in comment_patterns:
            self._add_comment_rule(pattern)

        # リテラル用ルールを追加
        literal_rules = [
            (r"'[^']*'", COLOR_ROYALBLUE), 
            (r"'(?:''|[^'])*'", COLOR_ROYALBLUE), 
            (r'"[^"]*"', COLOR_ROYALBLUE), 
            (r"\b\d+(?:\.\d+)?\b", COLOR_SLATEBLUE),
        ]
        for rule in literal_rules:
            pattern, color = rule
            self._add_literal_rule(pattern, color)

    def highlightBlock(self, text):
        # キーワード・リテラル・関数の色づけ
        self._apply_rules(text)
        # 複数行コメント対応
        self._highlight_multiline_comment(text)

    def _apply_rules(self, text):
        """ ルールを適用してテキストに色づけする"""
        for pattern, fmt in self.rules:
            for match in pattern.finditer(text):
                self.setFormat(
                    match.start(), 
                    match.end() - match.start(), 
                    fmt
                )

    def _highlight_multiline_comment(self, text):
        # 現在コメントブロック内かどうか取得
        fmt = QTextCharFormat()
        fmt.setForeground(QColor(COLOR_DARKGRAY))
        state = self.previousBlockState()
        start = 0
        # 現在ブロックコメント内
        if state == LexerState.BLOCK_COMMENT:
            end = text.find("*/")
            # 行内に`*/`なし -> 行全体を塗る
            if end == -1:
                self.setFormat(start, len(text), fmt)
                self.setCurrentBlockState(LexerState.BLOCK_COMMENT)
                return
            # 行内に`*/`あり -> 先頭から`*/`まで塗る
            self.setFormat(start, end + len("*/"), fmt)
            self.setCurrentBlockState(LexerState.NORMAL)
            return
        # 現在通常状態
        else:
            while True:
                start = text.find("/*", start)
                # 見つからない -> 塗る必要なし
                if start == -1:
                    self.setCurrentBlockState(LexerState.NORMAL)
                    break
                # ここへ来た時点で`/*`あり
                end = text.find("*/", start)
                # 行内に`*/`なし -> `/*`以降を塗ってフラグを立てる
                if end == -1:
                    self.setFormat(start, len(text) - start, fmt)
                    self.setCurrentBlockState(LexerState.BLOCK_COMMENT)
                    return
                self.setFormat(start, end - start + len("*/"), fmt)
                start = end + len("*/")

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

    def _add_function_rule(self, keyword: str):
        fmt = QTextCharFormat()
        fmt.setForeground(
            QColor(COLOR_DEEPPINK)
        )

        self.rules.append(
            (
                re.compile(
                    rf"\b{keyword}\s*(?=\()", 
                    re.IGNORECASE
                ), 
                fmt, 
            )
        )

    def _add_comment_rule(self, pattern: str):
        fmt = QTextCharFormat()
        fmt.setForeground(
            QColor(COLOR_FORESTGREEN)
        )
        self.rules.append(
            (
                re.compile(
                    pattern
                ), 
                fmt
            )
        )
    
    def _add_literal_rule(self, pattern: str, color: str):
        fmt = QTextCharFormat()
        fmt.setForeground(
            QColor(color)
        )
        self.rules.append(
            (
                re.compile(
                    pattern
                ), 
                fmt
            )
        )
