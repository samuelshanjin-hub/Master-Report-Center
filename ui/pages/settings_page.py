from PyQt6.QtWidgets import QWidget, QVBoxLayout, QLabel
class SettingsPage(QWidget):
    def __init__(self):
        super().__init__(); l=QVBoxLayout(self)
        t=QLabel("Settings"); t.setObjectName("PageTitle"); l.addWidget(t)
        x=QLabel("V2 Structure reserves this page for saved configurations, Excel engine preferences and output defaults.")
        x.setWordWrap(True); x.setObjectName("BodyText"); l.addWidget(x); l.addStretch()
