from typing import List

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, 
    QHBoxLayout, 
    QLabel, QListWidget, 
    QPushButton, 
    QTableWidget, QTableWidgetItem)
from PyQt6.QtCore import Qt

from config import TABLE_NAMES_FILE
from app.core.paths import get_base_dir
from app.models.query_result import QueryResult
from app.infrastructure.excel_runner import ExcelRunner
from app.infrastructure.json_loader import JSONLoader
from app.core.text_utils import truncate
from app.ui.widgets.sql_editor import SQLEditor
from app.ui.widgets.table_browser_dialog import TableBrowserDialog

class MainWindow(QMainWindow):
    def __init__(self):
        """ 
        MainWindowクラスのコンストラクタ

        .. note::
        - app/ui/main_window.py
        """
        super().__init__()

        # メイン・ウィンドウの設定
        # タイトル
        self.setWindowTitle("Portable SQL Client")
        # ウィンドウサイズ（固定）
        self.resize(800, 600)
        # SQL入力用エディタ
        self.sql_editor = SQLEditor()
        # テーブル一覧表示用リスト用のラベル
        table_list_label = QLabel("Table names")
        table_list_label.setStyleSheet("""
            background-color: #eee;
            color: #111;
            font-weight: bold;
            padding: 4px;
        """)
        # テーブル一覧表示用リスト
        self.table_list = QListWidget()
        # テーブル名リスト読み込み
        self.table_list.addItems(self._load_table_names())
        # SQL実行ボタン（サイズ固定）
        self.exec_button = QPushButton("クエリ実行！")
        self.exec_button.setFixedHeight(40)
        self.exec_button.setFixedWidth(80)
        # マウスホバーでポインタ変更
        self.exec_button.setCursor(
            Qt.CursorShape.PointingHandCursor
        )
        # マウスホバーで色を変更
        self.exec_button.setStyleSheet(
            """QPushButton:hover { background-color: #e0ffff; }""")
        # 結果表示テーブル
        self.result_table = QTableWidget()

        # UIパーツの配置
        # 中央に親ウィジェットを配置
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        # メインボックスレイアウトの作成
        layout = QVBoxLayout()
        central_widget.setLayout(layout)
        # 最上段のボックスレイアウト
        top_layout = QHBoxLayout()
        # ラベル + テーブル名リストのレイアウト
        table_list_laytout = QVBoxLayout()
        table_list_laytout.addWidget(table_list_label)
        table_list_laytout.addWidget(self.table_list)
        top_layout.addWidget(self.sql_editor, 2)
        top_layout.addLayout(table_list_laytout, 1)

        # テーブル名リストのアイテムにダブルクリックイベントをバインド
        self.table_list.itemDoubleClicked.connect(
            self._on_table_list_double_clicked
        )

        # ボタンのレイアウト
        button_layout = QHBoxLayout()
        # 〝伸びる空白（stretch）〟を追加（add）して右寄せにする
        button_layout.addStretch()
        button_layout.addWidget(self.exec_button)

        # テキストエリアを追加
        # 最上段のレイアウトをメインレイアウトに追加
        layout.addLayout(top_layout, 3)
        # layout.addWidget(self.sql_editor, 3)
        # 「クエリ実行！」ボタンを追加
        layout.addLayout(button_layout)

        # 結果表示テーブルを追加
        layout.addWidget(self.result_table, 2)

        # コマンドのバインド
        #   - 関数オブジェクトとして渡す
        self.exec_button.clicked.connect(
            self._on_exec_button_clicked
        )

    # Private method
    # テーブル名リストを読み込む
    def _load_table_names(self) -> List[str]:
        with open(TABLE_NAMES_FILE, encoding="utf-8") as f:
            # テキストファイルから1行ずつ取り出してリストにして返す
            return [
                line.strip() for line in f 
                # 空行、`#`で始まる行は無視
                if (line.strip() and not line.lstrip().startswith("#"))
            ]

    #  クリックイベントを受け取る
    def _on_exec_button_clicked(self):
        """ 「クエリ実行！」ボタンのクリックイベントの処理"""
        # 想定外の例外をキャッチ
        try:
            # エディタからクエリ取り出し
            query = self.sql_editor.toPlainText()
            # 踏み台Excelにクエリを投げる
            runner = ExcelRunner()
            # クエリ実行
            runner.execute(
                output_path=get_base_dir() / "temp" / "result.json", 
                query=query, 
                params=[], 
                timeout=30
            )
        except Exception as e:
            self._show_query_result(
                QueryResult.error(
                    title="クエリ実行失敗", 
                    message=f"on ExcelRunner.execute(): {str(e)}"
                )
            )
            return

        try:
            # JSON読み込み
            loader = JSONLoader()
            result = loader.load(
                get_base_dir() / "temp" / "result.json"
            )
        except Exception as e:
            self._show_query_result(
                QueryResult.error(
                    title="JSON読み込み失敗", 
                    message=f"on JSONLoader.load(): {str(e)}"
                )
            )
            return

        if result.is_error:
            self._show_query_result(result)
            return

        try:
            # 結果セット表示
            self._show_query_result(result)
        except Exception as e:
            self._show_query_result(
                QueryResult.error(
                    title="結果表示失敗", 
                    message=f"on MainWindow._show_query_result(): {str(e)}"
                )
            )
            return

    def _show_query_result(
            self, 
            result: QueryResult):
        """ 結果セットをテーブルに表示する"""
        MAX_COL_WIDTH = 200
        
        self.result_table.setColumnCount(
            result.column_count
        )
        self.result_table.setRowCount(
            result.row_count
        )
        # ヘッダを表示
        self.result_table.setHorizontalHeaderLabels(
            result.columns
        )
        # レコードを表示
        for row_index, row in enumerate(result.rows):
            for col_index, value in enumerate(row):
                # セルにはQTableWidgetItemインスタンスを闘魂注入
                text = str(value)
                item = QTableWidgetItem(text)
                item.setToolTip(text)
                self.result_table.setItem(
                    row_index, 
                    col_index, 
                    item    
                )

        # 内容に応じて列幅自動調整
        self.result_table.resizeColumnsToContents()
        # テーブルのヘッダの設定
        self.result_table.horizontalHeader().setStyleSheet(
            """
            QHeaderView::section {
                background-color: #333;
                color: #fff;
                font-weight: bold;
            }
            """
        )
        # 行の背景色の縞模様
        self.result_table.setAlternatingRowColors(True)
        # セル選択時に行全体をハイライト
        self.result_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )

        # 最大幅を制限
        for col in range(self.result_table.columnCount()):
            width = self.result_table.columnWidth(col)

            if width > MAX_COL_WIDTH:
                self.result_table.setColumnWidth(
                    col, 
                    MAX_COL_WIDTH
                )

    # テーブル情報表示ダイアログまわり
    def _on_table_list_double_clicked(self, item):
        # テーブル名を渡してTableBrowserDialogインスタンスを作成
        table_name = item.text()
        # キャッシュ or DBからテーブル情報取得
        data = self._load_table_info(table_name)
        dialog = TableBrowserDialog(table_name, data, self)
        # ダイアログ表示
        dialog.show()

    def _load_table_info(self, table_name: str) -> QueryResult:
        table_info_cache = get_base_dir() / "cache" / f"{table_name}.json"
        if not table_info_cache.exists():
            runner = ExcelRunner()
            runner.get_table_info(table_name)
            # 踏み台Excelでの処理後、JSONがない -> 通信失敗
            if not table_info_cache.exists():
                err_msg = \
                    "Excel側でI/Oエラーが発生した可能性があります。"
                return QueryResult(
                    columns=["( ´,_ゝ`)", "ち～ん（笑）"], 
                    rows=[
                        ["残念ｗ", "JSONが取得できなかったｗｗｗ"], 
                        ["原因はたぶん……", truncate(err_msg)]
                    ]
                )

        # cacheフォルダのテーブル情報読み込み
        loader = JSONLoader()
        return loader.load(table_info_cache)