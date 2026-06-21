from PyQt6.QtWidgets import QDialog

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
        print(data.columns)
        print(data.rows[:3])
        
        # ウィンドウサイズ
        self.resize(400, 600)