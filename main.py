import sys
from pathlib import Path

from PyQt6.QtWidgets import QApplication

from app.ui.main_window import MainWindow
from app.core.paths import get_base_dir
from app.infrastructure.excel_runner import ExcelRunner
from app.infrastructure.json_loader import JSONLoader

TEST_QUERY = """
SELECT
    pf.PrefecturalID
    , pf.PrefecturalName
FROM 
    Prefecturals pf
WHERE 
    pf.PrefecturalID IN (?, ?, ?)
;
"""

def main():
    app = QApplication(sys.argv)

    window = MainWindow()
    window.show()

    sys.exit(app.exec())

if __name__ == "__main__":
    main()
