# portable_sql_cloent README
## プロジェクトのフォルダ構成
```bash
portable_sql_client/
    ├ .venv/
    ├ main.py
    ├ app/
    │   ├ core/
    │   |    ├ paths.py
    │   |    └ text_utils.py
    │   ├ infrastructure/
    │   |    ├ excel_runner.py
    │   |    └ json_loader.py
    │   ├ models/
    │   |    └ query_result.py
    │   ├ services/
    │   └ ui/
    │        ├ widgets/
    │        |    ├ sql_editor.py
    │        |    ├ sql_highlighter.py
    │        |    └ table_browser.py
    │        └ main_window.py
    ├ assets/
    │   └ ba-90.ico
    ├ excel/
    │   ├ vba/
    │   └ QueryRunner.xlsm
    ├ logs/
    ├ temp/
    ├ tests/
    ├ .gitignore
    ├ config.py
    ├ main.py
    ├ readme.md # this file
    ├ requirements.txt
    ├ table_names.txt
    └ README.md

```
### `app`
- `ui/`
    - PyQt画面
    - 入力
    - 結果表示
    - 👉 SQLも知らない

- `services/`
    - 👉 〝ユースケース層〟
    - 例：
        - `execute_query()`
        - `run_sql()`
        - `get_result()`
    - 👉 UIとcoreの橋渡しだけ

- `core/`（重要）
    - 👉 アプリの中心ロジック
    - ここに入れるべき：
        - `QueryResult`
        - SQL実行フロー制御
        - エラー統一モデル
    - 👉 UIにもExcelにも依存しない

- `infrastructure/`
    - 👉 外部世界との接続
        - `ExcelRunner`（xlsm起動・呼び出し）
        - `JSONLoader`（temp読む）
        - File I/O
    - 👉 Excel依存は全部ここ

- `models/`（軽くでOK）
    - `QueryResult`モデル
    - 👉 ただし最初は薄くていい
