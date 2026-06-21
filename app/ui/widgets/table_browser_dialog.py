from PyQt6.QtWidgets import (
    QDialog, QTableWidget, QTableWidgetItem, 
    QVBoxLayout)

from app.models.query_result import QueryResult

class TableBrowserDialog(QDialog):
    def __init__(self, table_name: str, data: QueryResult, parent=None):
        """ テーブル名と表示データを渡してダイアログを初期化する
        
        :param table_name: テーブル名
        :type table_name: str
        :param data: テーブル情報オブジェクト
        :type data: QueryResult

        .. note::
        - app/ui/widgets/table_browser_dialog.py
        """
        super().__init__(parent)
        
        # ウィンドウのタイトルをセット
        self.setWindowTitle(f"Table infomation - {table_name}")
        
        # 取得したデータをウィンドウに表示
        self.table = QTableWidget()
        # 列ラベルのデータ
        self.table.setColumnCount(len(data.columns))
        self.table.setHorizontalHeaderLabels(data.columns)
        # 各行のデータ
        self.table.setRowCount(len(data.rows))
        # テーブルにデータを闘魂注入
        for row_no, row in enumerate(data.rows):
            for col_no, value in enumerate(row):
                text = str(value)
                item = QTableWidgetItem(text)
                # ツールチップ設定
                item.setToolTip(text)
                self.table.setItem(
                    row_no, 
                    col_no, 
                    item
                )
        # 表示内容に応じて列幅調整
        self.table.resizeColumnsToContents()
        # ヘッダの表示設定
        self.table.horizontalHeader().setStyleSheet(
            """
            QHeaderView::section {
                background-color: #333;
                color: #fff;
                font-weight: bold;
            }
            """
        )

        # テーブルをウィンドウにレイアウトする
        layout = QVBoxLayout()
        layout.addWidget(self.table)
        self.setLayout(layout)
        
        # ウィンドウサイズ
        self.resize(400, 600)