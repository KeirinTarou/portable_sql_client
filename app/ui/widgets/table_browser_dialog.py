from PyQt6.QtWidgets import QDialog

from app.core.paths import get_base_dir
from app.infrastructure.excel_runner import ExcelRunner
from app.infrastructure.json_loader import JSONLoader

class TableBrowserDialog(QDialog):
    def __init__(self, table_name: str, parent=None):
        super().__init__(parent)
        
        # ウィンドウのタイトルをセット
        self.setWindowTitle(f"Table infomation - {table_name}")

        # テーブル情報のキャッシュがないときは取得 -> cacheフォルダに保存
        table_info_cache = get_base_dir() / "cache" / f"{table_name}.json"
        if not table_info_cache.exists():
            runner = ExcelRunner()
            runner.get_table_info(table_name)
        # cacheフォルダのテーブル情報読み込み
        loader = JSONLoader()
        data = loader.load(table_info_cache)
        
        # 取得したデータをウィンドウに表示
        print(data.columns)
        print(data.rows[:3])
        
        # ウィンドウサイズ
        self.resize(400, 600)