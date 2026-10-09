from PyQt6.QtWidgets import QWidget, QVBoxLayout, QLabel
class HistoryPage(QWidget):
    def __init__(self):
        super().__init__(); l=QVBoxLayout(self)
        t=QLabel("History"); t.setObjectName("PageTitle"); l.addWidget(t)
        x=QLabel("Job history is intentionally separated from module logic and can be connected to HistoryService in the next production iteration.")
        x.setWordWrap(True); x.setObjectName("BodyText"); l.addWidget(x); l.addStretch()
