# app/ui/widgets/sql_highlighter.py _analize_lines()のテスト

from app.ui.widgets.sql_highlighter import (
    _analyze_line, LexerState)

# TODO:
#   - 通常時: 
#       ✅- `/*`がない -> NORMAL, 空のリスト
#       ✅- `/*`あり -> BLOCK_COMMENT, コメント位置情報リスト
#       ✅- `/*`、`*/`ともにあり -> NORMAL, コメント位置情報リスト
#       ✅- `/* ... */`が1行に2箇所 -> NORMAL, コメント位置情報リスト

# 通常時
def test_analyze_lines_no_comment():
    """ 通常時: `/*`がない -> NORMAL, 空のリスト"""
    text = "IRON MAIDEN SEVENTH SON OF THE SEVENTH SON"
    next_state, regions = _analyze_line(text, LexerState.NORMAL)
    assert next_state == LexerState.NORMAL
    assert regions == []

def test_analyze_lines_comment_start():
    """ 通常時: `/*`あり -> BLOCK_COMMENT, コメント位置情報リスト"""
    text = "IRON MAIDEN /*KILLERS"
    next_state, regions = _analyze_line(text, LexerState.NORMAL)
    assert next_state == LexerState.BLOCK_COMMENT
    assert regions == [[12, 21]]

def test_analyze_lines_comment_start_and_end():
    """ 通常時: `/*`、`*/`ともにあり -> NORMAL, コメント位置情報リスト"""
    text = "IRON MAIDEN THE /*NUMBER*/ OF THE BEAST"
    next_state, regions = _analyze_line(text, LexerState.NORMAL)
    assert next_state == LexerState.NORMAL
    assert regions == [[16, 26]]

def test_analyze_lines_two_separate_comments():
    """ 通常時: `/* ... */`が1行に2箇所 -> NORMAL, コメント位置情報リスト"""
    text = "IRON /*MAIDEN*/ SOMEWHERE /*IN*/ TIME"
    next_state, regions = _analyze_line(text, LexerState.NORMAL)
    assert next_state == LexerState.NORMAL
    assert regions == [[5, 15], [26, 32]]
