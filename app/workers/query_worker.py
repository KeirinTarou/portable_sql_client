from PyQt6.QtCore import (QThread, pyqtSignal)

from app.core.paths import get_base_dir
from app.infrastructure.excel_runner import ExcelRunner
from app.infrastructure.json_loader import JSONLoader
from app.models.query_result import QueryResult

class QueryWorker(QThread):
    def __init__(
            self, query: str, parent=None):
        """ 第1引数で実行するクエリを渡す"""
        super().__init__(parent)
        self.query = query

    def run(self):
        # 想定外の例外をキャッチ
        try:
            # 踏み台Excelにクエリを投げる
            runner = ExcelRunner()
            # クエリ実行
            runner.execute(
                output_path=get_base_dir() / "temp" / "result.json", 
                query=self.query, 
                params=[], 
                timeout=30
            )
        except Exception as e:
            result = QueryResult.error(
                title="クエリ実行失敗", 
                message=f"on ExcelRunner.execute(): {str(e)}"
            )
            print(result)
            # ここで処理を終える必要がある
            return

        try:
            # JSON読み込み
            loader = JSONLoader()
            result = loader.load(
                get_base_dir() / "temp" / "result.json"
            )
        except Exception as e:
            result = QueryResult.error(
                title="JSON読み込み失敗", 
                message=f"on JSONLoader.load(): {str(e)}"
            )
            print(result)
            # ここで処理を終える必要がある
            return
        
        print(result)
