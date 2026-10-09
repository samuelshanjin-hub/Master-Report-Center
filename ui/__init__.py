PUBLIC_STYLE = r"""
QMainWindow, QWidget {
    background: #F6F8F7;
    color: #17231D;
    font-family: 'Segoe UI';
    font-size: 13px;
}
#TopBar { background: #FFFFFF; border-bottom: 1px solid #DCE6DF; }
#Sidebar { background: #FFFFFF; border-right: 1px solid #DCE6DF; }
#ContentArea { background: #F6F8F7; }
#HomeScroll { background:#F6F8F7; border:none; }
#HomeScroll > QWidget > QWidget { background:#F6F8F7; }
#HomeContent { background:#F6F8F7; }
QScrollBar:vertical { background:#F1F5F2; width:11px; margin:0; }
QScrollBar::handle:vertical { background:#BFCFC4; border-radius:5px; min-height:28px; }
QScrollBar::handle:vertical:hover { background:#9FB6A7; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height:0; }
QScrollBar:horizontal { background:#F1F5F2; height:11px; margin:0; }
QScrollBar::handle:horizontal { background:#BFCFC4; border-radius:5px; min-width:28px; }
QScrollBar::handle:horizontal:hover { background:#9FB6A7; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width:0; }
#Card { background: #FFFFFF; border: 1px solid #D9E4DD; border-radius: 12px; }
#UtilityPanel { background: #FBFDFC; border: 1px solid #CFE1D4; border-radius: 10px; }
#Panel { background: #FFFFFF; border: 1px solid #D9E4DD; border-radius: 10px; }
#InfoPanel { background: #FFFFFF; border: 1px solid #D9E4DD; border-radius: 10px; }
#Divider { color: #DCE6DF; }
QLabel { background: transparent; }
QLabel#HeaderBrand { color:#173D2A; font-size:14px; font-weight:600; }
QLabel#AppTitle { color: #143D2A; font-size: 28px; font-weight: 700; }
QLabel#HomeSubtitle { color:#66776E; font-size:15px; font-weight:600; letter-spacing:0.4px; }
QLabel#HomeWelcome { color:#17612F; font-size:17px; font-weight:700; }
QLabel#PageTitle { color: #102B20; font-size: 25px; font-weight: 700; }
QLabel#PageSubtitle { color: #66776E; font-size: 13px; }
QLabel#SectionTitle { color: #124E2B; font-size: 15px; font-weight: 700; }
QLabel#Muted { color: #66776E; }
QLabel#Success { color:#187A32; font-weight:600; }
QPushButton { background: #FFFFFF; border: 1px solid #CAD8CF; border-radius: 7px; padding: 8px 14px; }
QPushButton:hover { background: #EFF6F1; border-color: #79A88B; }
QPushButton#Secondary, QPushButton#SecondaryButton { background:#EAF4EC; color:#155B2D; border:1px solid #B8D6C0; font-weight:600; }
QPushButton#Secondary:hover, QPushButton#SecondaryButton:hover { background:#DDEEDF; }
QPushButton#Info { background:#EAF2FB; color:#0F5EA5; border:1px solid #B8CEE8; font-weight:600; }
QPushButton#Info:hover { background:#DCEAF8; border-color:#8FB2D9; }
QPushButton#Primary, QPushButton#PrimaryButton { background:#0F6CBD; color:white; border:none; font-size:14px; font-weight:700; }
QPushButton#Primary:hover, QPushButton#PrimaryButton:hover { background:#0D5EA5; }
QPushButton#Green { background:#148A32; color:white; border:none; font-size:14px; font-weight:700; }
QPushButton#Green:hover { background:#117328; }
QPushButton#Purple { background:#5B2296; color:white; border:none; font-size:14px; font-weight:700; }
QPushButton#Purple:hover { background:#4D1D80; }
QPushButton#Teal { background:#0F766E; color:white; border:none; font-size:14px; font-weight:700; }
QPushButton#Teal:hover { background:#0B625C; }
QPushButton#Nav { text-align:left; border:none; background:transparent; padding:10px 14px; font-size:14px; }
QPushButton#Nav:hover { background:#EEF5F0; }
QPushButton#NavActive { text-align:left; border:1px solid #CDE0D2; border-radius:7px; background:#EAF4EC; color:#155B2D; font-weight:700; padding:10px 14px; }
QPushButton#UtilityAction { text-align:left; border:none; background:transparent; color:#155B2D; font-weight:600; padding:7px 8px; }
QPushButton#UtilityAction:hover { background:#EEF5F0; }
QLineEdit, QComboBox { background:white; border:1px solid #C9D6CE; border-radius:6px; padding:8px; min-height:22px; }
QTextEdit, QPlainTextEdit { background:white; border:1px solid #CDD9D1; border-radius:8px; }
QProgressBar { border:1px solid #D4DED8; border-radius:6px; background:#E8EFEA; text-align:center; min-height:18px; }
QProgressBar::chunk { background:#0F6CBD; border-radius:5px; }
QGroupBox { background:white; border:1px solid #D8E2DC; border-radius:9px; margin-top:12px; font-weight:700; }
QGroupBox::title { subcontrol-origin:margin; left:12px; padding:0 5px; color:#173D2A; }
QLabel#FilenamePreview {
    background:#F7FAF8;
    color:#173D2A;
    border:1px solid #D7E3DB;
    border-radius:6px;
    padding:7px 9px;
    font-family:'Segoe UI';
}
QCheckBox { spacing:8px; }
QStatusBar { background:#FFFFFF; color:#66776E; border-top:1px solid #DCE6DF; }

#LicenseHomePanel { background:#FFFFFF; border:1px solid #D6E0DA; border-radius:11px; }
QLabel#LicenseStatus { color:#173F32; font-size:13px; font-weight:800; }
QPushButton#LicenseBuy { background:#1F5E46; color:#FFFFFF; border:none; font-weight:700; min-width:105px; }
QPushButton#LicenseBuy:hover { background:#174B38; }
QPushButton#LicenseAction { background:#F4F7F5; color:#28483A; border:1px solid #C8D5CD; font-weight:700; }
QPushButton#LicenseAction:hover { background:#EAF1ED; border-color:#9FB8A9; }
"""

