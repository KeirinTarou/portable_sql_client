from PyQt6.QtCore import (QThread, pyqtSignal)

from app.core.paths import get_base_dir
from app.infrastructure.excel_runner import ExcelRunner
from app.infrastructure.json_loader import JSONLoader
from app.models.query_result import QueryResult

class TableInfoWorker(QThread):
    result_ready = pyqtSignal(str, QueryResult)

    def __init__(self, table_name: str):
        super().__init__()
        self.table_name = table_name

    def run(self):
        cache_path = get_base_dir() / "cache" / f"{self.table_name}.json"

        # キャッシュしたテーブル情報JSONなし -> 踏み台Excelから取得
        if not cache_path.exists():
            runner = ExcelRunner()
            runner.get_table_info(self.table_name)

        # ここに来た時点でテーブル情報JSONなし -> 異常事態 -> エラー表示
        if not cache_path.exists():
            self.result_ready.emit(
                self.table_name, 
                QueryResult.error(
                    title="踏み台Excelのエラー", 
                    message="JSONエクスポート時のI/Oエラーの可能性あり。"
                )
            )
        
        # テーブル情報JSON読み込み
        loader = JSONLoader()
        result = loader.load(cache_path)
        # シグナル発行
        self.result_ready.emit(self.table_name, result)