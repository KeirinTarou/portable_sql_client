# portable_sql_cloent README
## プロジェクトのフォルダ構成
```bash
portable_sql_client/
├ .venv/
├ main.py
├ app/
│   ├ core/
│   ├ domain/
│   ├ infrastructure/
│   ├ services/
│   └ ui/
├ tests/
├ excel/
│   └ QueryRunner.xlsm
├ temp/
├ logs/
├ requirements.txt
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
        - ExcelRunner（xlsm起動・呼び出し）
        - JSONLoader（temp読む）
        - File IO
    - 👉 Excel依存は全部ここ

- `domain/`（軽くでOK）
    - Resultモデル
    - Status（SUCCESS / ERROR）
    - 👉 ただし最初は薄くていい
