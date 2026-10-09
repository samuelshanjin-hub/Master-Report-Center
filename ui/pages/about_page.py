from PyQt6.QtWidgets import QWidget, QVBoxLayout, QLabel
from version import BRAND, STRUCTURE_NAME, VERSION

class AboutPage(QWidget):
    def __init__(self):
        super().__init__()
        l = QVBoxLayout(self)
        t = QLabel("About")
        t.setObjectName("PageTitle")
        l.addWidget(t)
        text = f"{BRAND}\n\n{STRUCTURE_NAME}\nVersion {VERSION}\n\nA modular desktop framework for report utilities."
        b = QLabel(text)
        b.setObjectName("BodyText")
        b.setWordWrap(True)
        l.addWidget(b)
        l.addStretch()
