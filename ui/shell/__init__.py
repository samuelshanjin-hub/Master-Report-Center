from PyQt6.QtCore import pyqtSignal
from PyQt6.QtWidgets import QFrame, QVBoxLayout, QPushButton, QLabel

class Sidebar(QFrame):
    navigate = pyqtSignal(str)
    def __init__(self):
        super().__init__()
        self.setObjectName("Sidebar")
        self.setFixedWidth(190)
        layout = QVBoxLayout(self)
        brand = QLabel("GFS\nMASTER REPORT\nCENTER")
        brand.setObjectName("SidebarBrand")
        layout.addWidget(brand)
        for key, label in [("home","Home"),("history","History"),("settings","Settings"),("about","About")]:
            btn = QPushButton(label)
            btn.setObjectName("NavButton")
            btn.clicked.connect(lambda checked=False, k=key: self.navigate.emit(k))
            layout.addWidget(btn)
        layout.addStretch()

