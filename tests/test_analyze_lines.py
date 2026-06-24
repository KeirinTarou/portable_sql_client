# app/ui/widgets/sql_highlighter.py _analize_lines()のテスト

from app.ui.widgets.sql_highlighter import (
    _analyze_line, LexerState)

# TODO:
#   - 通常時: 
#       ✅- `/*`がない -> NORMAL, 空のリスト
#       ✅- `/*`あり -> BLOCK_COMMENT, コメント位置情報リスト
#       ✅- `/*`、`*/`ともにあり -> NORMAL, コメント位置情報リスト
#       ✅- `/* ... */`が1行に2箇所 -> NORMAL, コメント位置情報リスト
#       ✅- `/* ... */ ... /* ...` -> BLOCK_COMMENT, コメント位置情報リスト
#       ✅- 通常時: `/* ... /* ...` -> BLOCK_COMMENT, [`/*`先頭, 行末]
#       - `--`の後ろの`/*`は無視
#       - `'/*'`は無視
#       - `"/*"`は無視
#       - `/*`が`*/`で閉じられる前の`/*`は無視
#       - 最初に出てくる`*/`は無視
#       - `'*/'`は無視
#       - `"*/"`は無視
#   - ブロックコメント内
#       ✅- `*/`がない -> BLOCK_COMMENT, [行の先頭, 行の末尾]
#       ✅- `*/`あり -> NORMAL, [行の先頭, `*/`の末尾]
#       ✅- `*/`の後に`/*`あり 
#           -> BLOCK_COMMENT, [[行の先頭, `*/`の末尾], [`/*`の先頭, 行の末尾]] 
#       ✅- `*/`の後に`/* ... */`あり 
#           -> NORMAL, [[行の先頭, `*/`の末尾], [`/*`の先頭, `*/`の末尾]] 

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

def test_analyze_lines_closed_and_unclosed_comment():
    """ 通常時: `/* ... */ ... /* ...` -> BLOCK_COMMENT, コメント位置情報リスト"""
    text = "IRON /*MAIDEN*/ NO PRAYER FOR /*THE DYING"
    next_state, regions = _analyze_line(text, LexerState.NORMAL)
    assert next_state == LexerState.BLOCK_COMMENT
    assert regions == [[5, 15], [30, 41]]

# ブロックコメント内
def test_analyze_lines_no_closed():
    """ ブロックコメント内: `*/`なし -> BLOCK_COMMENT, [行先頭, 行末尾]"""
    text = "IRON MAIDEN POWERSLAVE"
    next_state, regions = _analyze_line(text, LexerState.BLOCK_COMMENT)
    assert next_state == LexerState.BLOCK_COMMENT
    assert regions == [[0, 22]]

def test_analyze_lines_closed():
    """ ブロックコメント内: `*/`あり -> NORMAL, [行先頭, `*/`末尾]"""
    text = "IRON MAIDEN */POWERSLAVE"
    next_state, regions = _analyze_line(text, LexerState.BLOCK_COMMENT)
    assert next_state == LexerState.NORMAL
    assert regions == [[0, 14]]

def test_analyze_lines_closed_and_opened():
    """ ブロックコメント内: `... */ ... /* ...` 
        -> BLOCK_COMMENT, [[行先頭, `*/`末尾], [`/*`先頭, 行末尾]]"""
    text = "IRON*/ MAIDEN /*FEAR OF THE DARK"
    next_state, regions = _analyze_line(text, LexerState.BLOCK_COMMENT)
    assert next_state == LexerState.BLOCK_COMMENT
    assert regions == [[0, 6],[14, 32]]

def test_analyze_lines_closed_and_recommented():
    """ ブロックコメント内: `... */ ... /* ... */ ...`
        -> NORMAL, [[行先頭, `*/`末尾], [`/*`先頭, `*/`末尾]]"""
    text = "IRON*/ MAIDEN /*FEAR OF*/ THE DARK"
    next_state, regions = _analyze_line(text, LexerState.BLOCK_COMMENT)
    assert next_state == LexerState.NORMAL
    assert regions == [[0, 6], [14, 25]]

# 無視するトークン
def test_analyze_lines_ignore_opener_before_closed():
    """ 通常時: `... /* ... /* ...` -> BLOCK_COMMENT, [`/*`先頭, 行末]
        ブロックコメント内: `... /* ... */ ...` -> NORMAL, [行先頭, `*/`末尾]"""
    text = "IROM /*MAIDEN /*KILLERS"
    # 通常時
    next_state, regions = _analyze_line(text, LexerState.NORMAL)
    assert next_state == LexerState.BLOCK_COMMENT
    assert regions == [[5, 23]]
    # コメントブロック内
    text = "IROM /*MAIDEN */KILLERS"
    next_state, regions = _analyze_line(text, LexerState.BLOCK_COMMENT)
    assert next_state == LexerState.NORMAL
    assert regions == [[0, 16]]