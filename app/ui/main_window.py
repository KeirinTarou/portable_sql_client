from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, 
    QHBoxLayout, 
    QPushButton, 
    QTableWidget, QTableWidgetItem)
from PyQt6.QtCore import Qt

from app.core.paths import get_base_dir
from app.models.query_result import QueryResult
from app.infrastructure.excel_runner import ExcelRunner
from app.infrastructure.json_loader import JSONLoader
from app.core.text_utils import truncate
from app.ui.widgets.sql_editor import SQLEditor

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
        # SQL実行ボタン（サイズ固定）
        self.exec_button = QPushButton("クエリ実行！")
        self.exec_button.setFixedHeight(40)
        self.exec_button.setFixedWidth(80)
        # マウスホバーでポインタ変更
        self.exec_button.setCursor(
            Qt.CursorShape.PointingHandCursor
        )
        # 結果表示テーブル
        self.result_table = QTableWidget()

        # UIパーツの配置
        # 中央に親ウィジェットを配置
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        # ボックスレイアウトの作成
        layout = QVBoxLayout()
        central_widget.setLayout(layout)
        # ボタンのレイアウト
        button_layout = QHBoxLayout()
        # 〝伸びる空白（stretch）〟を追加（add）して右寄せにする
        button_layout.addStretch()
        button_layout.addWidget(self.exec_button)

        # テキストエリアを追加
        layout.addWidget(self.sql_editor, 3)
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
            self._show_error(f"on ExcelRunner.execute(): {str(e)}")
            return

        try:
            # JSON読み込み
            loader = JSONLoader()
            result = loader.load(
                get_base_dir() / "temp" / "result.json"
            )
        except Exception as e:
            self._show_error(f"on JSONLoader.load(): {str(e)}")
            return

        if result.is_error:
            self._show_error(result.error_message)
            return

        try:
            # 結果セット表示
            self._show_query_result(result)
        except Exception as e:
            self._show_error(f"on MainWindwo._show_query_result(): {str(e)}")
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

    def _show_error(self, message: str):
        """ エラーメッセージを出力"""
        result = QueryResult(
            columns=["( ´,_ゝ`)", "ち～ん（笑）"], 
            rows=[
                ["残念ｗ", "レコードセットが返らなかったｗｗｗ"], 
                ["原因はたぶん……", truncate(message)]
            ]
        )
        self._show_query_result(result)
