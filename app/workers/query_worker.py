from PyQt6.QtCore import (QThread, pyqtSignal)

class QueryWorker(QThread):
    def run(self):
        print("Hello, from worker!!")
        