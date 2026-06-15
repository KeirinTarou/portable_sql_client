from pathlib import Path
from typing import List, Any

import win32com.client
import pythoncom

from app.core.paths import get_base_dir
from config import EXCEL_FILE_PATH
EXCEL_PATH = get_base_dir() / EXCEL_FILE_PATH

class ExcelRunner:
    def execute(
        self, 
        output_path: Path, 
        query: str, 
        params: List[Any], 
        timeout: int = 60) -> None:
        """ 
        踏み台Excelにクエリを投げてJSONを保存する

        :param output_path: 結果セットJSONのパス
        :type output_path: Path
        :param query: クエリ文字列
        :type query: str
        :param params: クエリに渡すパラメータ
        :type params: List[Any]
        :param timeout: タイムアウト秒数
        :type timeout: int = 60
        :return: 返り値なし
        :rtype: None
        .. note::
            - app/infrastructure/excel_runner.py
        """
        if params is None:
            params = []

        pythoncom.CoInitialize()

        excel = None
        wb = None

        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.visible = False
            wb = excel.Workbooks.Open(str(EXCEL_PATH))
            # 踏み台Excelのマクロに引数を渡して実行
            excel.Run(
                "'QueryRunner.xlsm'!Entry.ExecQuery", 
                str(output_path), 
                query, 
                params, 
                timeout
            )
        finally:
            if wb is not None:
                wb.Close(SaveChanges=False)
            if excel is not None:
                excel.Quit()
            # あとかたづけ
            pythoncom.CoUninitialize()
