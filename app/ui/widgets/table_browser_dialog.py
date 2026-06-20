from PyQt6.QtWidgets import QDialog

class TableBrowserDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("Table infomation")
        self.resize(400, 600)