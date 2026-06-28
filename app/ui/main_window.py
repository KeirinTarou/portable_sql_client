from typing import List

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, 
    QHBoxLayout, 
    QLabel, QListWidget, 
    QPushButton, 
    QTableWidget, QTableWidgetItem, 
    QApplication)
from PyQt6.QtCore import Qt

from config import TABLE_NAMES_FILE
from app.core.paths import get_base_dir
from app.models.query_result import QueryResult
from app.infrastructure.excel_runner import ExcelRunner
from app.infrastructure.json_loader import JSONLoader
from app.ui.widgets.sql_editor import SQLEditor
from app.ui.widgets.table_browser_dialog import TableBrowserDialog
from app.workers.query_worker import QueryWorker
from app.workers.table_info_worker import TableInfoWorker

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
        # 踏み台Excel別スレッド実行用ワーカー
        self.worker: QueryWorker | None = None
        # テーブル情報取得用ワーカー
        self.table_info_worker: TableInfoWorker | None = None
        # テーブル情報表示用ダイアログ
        self.dialogs: List[TableBrowserDialog] = []
        # クエリ実行ステータス
        self._is_running = False 
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
    def _set_running_state(self, running: bool):
        self._is_running = running
        # ボタンとテーブル一覧 -> 実行中無効
        self.exec_button.setEnabled(not running)
        self.table_list.setEnabled(not running)

        # エディタは実行中もReadOnlyにしない
        #   - 実行中ReadOnlyにしたくなったら下記コメントを解除
        # self.sql_editor.setReadOnly(running)

        # ボタンのラベルとマウスポインタの切り替え
        if running:
            self.exec_button.setText("実行中……")
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        else:
            self.exec_button.setText("クエリ実行！")
            QApplication.restoreOverrideCursor()

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
        # エディタのクエリを取り出す
        query = self.sql_editor.toPlainText()
        # 踏み台Excel用ワーカーを作成
        #   - すでに作成済みだったらreturn（何もしない）
        if self._is_running:
            return
        self.worker = QueryWorker(query)
        # result_readyシグナルに_show_query_result()をバインド
        #   - QueryWorkerでemit()を実行した時点で発火
        self.worker.result_ready.connect(
            self._show_query_result
        )
        # QueryWorkerの仕事が終わったことを知らせるシグナルを登録
        self.worker.finished.connect(
            self._on_query_finished
        )

        # UIを実行中状態に切り替える
        self._set_running_state(True)

        # QueryWorkerの仕事を始める
        #   - QueryResult準備作業開始
        #   - 準備ができたらresult_readyシグナルを発信
        self.worker.start()

    def _on_query_finished(self):
        # UIの状態を元に戻す
        self._set_running_state(False)

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
        # テーブル情報を取得するワーカーを作成
        worker = TableInfoWorker(table_name)
        # result_readyに_open_table_dialog()をバインド
        worker.result_ready.connect(self._open_table_info_dialog)
        # QueryWorkerの仕事が終わったことを知らせるシグナルを登録
        worker.finished.connect(
            self._on_query_finished
        )
        # UIを実行中状態に切り替える
        self._set_running_state(True)

        # ここでtable_info_worker属性にセット
        self.table_info_worker = worker
        # TableInfoWorker起動
        worker.start()

    def _open_table_info_dialog(self, table_name: str, result: QueryResult):
        # テーブル情報ダイアログを開く
        dialog = TableBrowserDialog(table_name, result, self)
        # dialogs属性に追加
        self.dialogs.append(dialog)
        # ダイアログを閉じたらリストから削除するよう予約
        # destroyedシグナル: Qtオブジェクト破棄通知
        #   - シグナルの情報を使わずともローカル変数dialogで削除対象は特定できる
        #   - よって、直接dialogを渡して削除する
        dialog.destroyed.connect(
            lambda _, d=dialog: self.dialogs.remove(d)
            if d in self.dialogs else None)

        # 同時に開けるのは最大5つまでにする
        if len(self.dialogs) > 5:
            d = self.dialogs.pop(0)
            d.close()

        dialog.show()
