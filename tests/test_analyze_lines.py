# app/ui/widgets/sql_highlighter.py _analize_lines()のテスト

from app.ui.widgets.sql_highlighter import (
    _analyze_line, LexerState)

# TODO:
#   - 通常時: 
#       ✅- `/*`がない -> NORMAL, 空のリスト

# 通常時
def test_analyze_lines_no_comment():
    """ 通常時: `/*`がない -> NORMAL, 空のリスト"""
    text = "IRON MAIDEN SEVENTH SON OF THE SEVENTH SON"
    next_state, regions = _analyze_line(text, LexerState.NORMAL)
    assert next_state == LexerState.NORMAL
    assert regions == []

