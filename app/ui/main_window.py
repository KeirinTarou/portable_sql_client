from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QPlainTextEdit)

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

        # UIパーツの配置
        # 中央に親ウィジェットを配置
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        # ボックスレイアウトの作成
        layout = QVBoxLayout()
        central_widget.setLayout(layout)

        # テキストエリアを追加
        layout.addWidget(self.sql_editor)

