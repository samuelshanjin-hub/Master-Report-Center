from datetime import datetime
from PyQt6.QtWidgets import QPlainTextEdit

class LogView(QPlainTextEdit):
    def __init__(self):
        super().__init__()
        self.setReadOnly(True)
        self.setMaximumBlockCount(2000)
    def add(self, message):
        ts = datetime.now().strftime("%H:%M:%S")
        self.appendPlainText(f"{ts}  {message}")
