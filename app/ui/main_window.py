from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, 
    QHBoxLayout, 
    QPlainTextEdit, QPushButton, 
    QTableWidget, QTableWidgetItem)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # メイン・ウィンドウの設定
        # タイトル
        self.setWindowTitle("Portable SQL Client")
        # ウィンドウサイズ（固定）
        self.resize(800, 600)
        # SQL入力用エディタ
        self.sql_editor = QPlainTextEdit()
        # SQL実行ボタン（サイズ固定）
        self.exec_button = QPushButton("クエリ実行！")
        self.exec_button.setFixedHeight(40)
        self.exec_button.setFixedWidth(80)
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
        self.result_table.setColumnCount(2)
        self.result_table.setRowCount(3)
