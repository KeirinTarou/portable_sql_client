from PyQt6.QtWidgets import QDialog

from app.ui.main_window import MainWindow

class TableBrowserDialog(QDialog):
    def __init__(self, table_name: str, parent: MainWindow=None):
        super().__init__(parent)

        self.setWindowTitle(f"Table infomation - {table_name}")
        self.resize(400, 600)