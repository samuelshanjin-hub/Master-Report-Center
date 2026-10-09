from PyQt6.QtWidgets import QHBoxLayout, QLabel, QLineEdit, QPushButton, QFileDialog

def add_file_picker(layout, label, mode="file", save=False):
    row=QHBoxLayout(); lab=QLabel(label); lab.setFixedWidth(160); edit=QLineEdit(); btn=QPushButton("Browse"); btn.setObjectName("SecondaryButton")
    def browse():
        if mode=="folder": value=QFileDialog.getExistingDirectory(None, label)
        elif save: value,_=QFileDialog.getSaveFileName(None, label, "", "Excel Workbook (*.xlsx)")
        else: value,_=QFileDialog.getOpenFileName(None, label, "", "Excel Workbook (*.xlsx)")
        if value: edit.setText(value)
    btn.clicked.connect(browse); row.addWidget(lab); row.addWidget(edit,1); row.addWidget(btn); layout.addLayout(row); return edit

