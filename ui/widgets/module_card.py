from PyQt6.QtCore import pyqtSignal
from PyQt6.QtWidgets import QFrame, QVBoxLayout, QLabel, QPushButton

class ModuleCard(QFrame):
    opened = pyqtSignal(str)
    def __init__(self, info):
        super().__init__()
        self.setObjectName("ModuleCard")
        layout = QVBoxLayout(self)
        icon = QLabel(info.icon_text)
        icon.setObjectName("ModuleIconText")
        title = QLabel(info.title)
        title.setObjectName("ModuleTitle")
        subtitle = QLabel(info.subtitle)
        subtitle.setObjectName("Muted")
        desc = QLabel(info.description)
        desc.setWordWrap(True)
        desc.setObjectName("BodyText")
        btn = QPushButton("Open Module")
        btn.setObjectName("PrimaryButton")
        btn.clicked.connect(lambda: self.opened.emit(info.module_id))
        for w in [icon, title, subtitle, desc, btn]: layout.addWidget(w)
        layout.addStretch()
