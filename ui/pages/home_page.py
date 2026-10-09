from pathlib import Path
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap
from PyQt6.QtWidgets import QWidget, QFrame, QLabel, QPushButton, QVBoxLayout, QHBoxLayout, QGridLayout, QGroupBox, QScrollArea, QSizePolicy
from version import VERSION

MODULE_PRESENTATION = {
    "split_reports": ("SPLIT REPORTS", "Excel Report Splitter", "Split one report\ninto multiple reports", "split_reports_logo.png", "Primary", "#164A8A"),
    "consolidate_reports": ("CONSOLIDATE REPORTS", "Excel Report Merger", "Combine multiple reports\ninto one report", "consolidate_reports_logo.png", "Purple", "#5B2A91"),
    "synchronize_reports": ("SYNCHRONIZE REPORTS", "Keep Reports in Sync", "Compare and synchronize\ndata between reports", "synchronize_reports_logo.png", "Green", "#21833A"),
    "lookup_data": ("LOOKUP DATA", "Data Lookup & Enrichment", "Find and retrieve data\nquickly and accurately", "lookup_data_logo.png", "Teal", "#0F766E"),
}

class ModuleCard(QFrame):
    def __init__(self, module_id, title, subtitle, description, image, button_style, title_color, command):
        super().__init__()
        self.setObjectName("Card")
        self.setMinimumHeight(260)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16,10,16,12)
        layout.setSpacing(4)

        pic = QLabel()
        pic.setMinimumHeight(105)
        pic.setMaximumHeight(135)
        pic.setAlignment(Qt.AlignmentFlag.AlignCenter)
        if Path(image).exists():
            pm = QPixmap(str(image))
            pic.setPixmap(pm.scaled(150,125,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation))

        title_lbl = QLabel(title)
        title_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_lbl.setStyleSheet(f"font-size:16px;font-weight:700;color:{title_color};")

        sub = QLabel(subtitle)
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sub.setStyleSheet("font-size:13px;font-weight:600;color:#222;")

        desc = QLabel(description)
        desc.setWordWrap(True)
        desc.setAlignment(Qt.AlignmentFlag.AlignCenter)
        desc.setStyleSheet("font-size:13px;color:#444;")

        btn = QPushButton("Open Module")
        btn.setObjectName(button_style)
        btn.clicked.connect(lambda: command(module_id))

        layout.addWidget(pic)
        layout.addWidget(title_lbl)
        layout.addWidget(sub)
        layout.addWidget(desc)
        layout.addStretch(1)
        layout.addWidget(btn)


class HomePage(QWidget):
    def __init__(self, app, modules):
        super().__init__()
        self.app = app
        self.modules = modules
        self.assets = Path(__file__).resolve().parents[2]/"assets"
        self._card_columns = None
        self.module_cards = []

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0,0,0,0)
        outer.setSpacing(0)

        self.scroll = QScrollArea()
        self.scroll.setObjectName("HomeScroll")
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        outer.addWidget(self.scroll)

        self.content = QWidget()
        self.content.setObjectName("HomeContent")
        self.scroll.setWidget(self.content)

        root = QVBoxLayout(self.content)
        root.setContentsMargins(28,12,28,14)
        root.setSpacing(8)

        brand = QLabel()
        brand.setMinimumHeight(90)
        brand.setMaximumHeight(145)
        logo = self.assets/"golden_fir_wordmark.png"
        if logo.exists():
            pm = QPixmap(str(logo))
            brand.setPixmap(pm.scaled(520,130,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation))
        brand.setAlignment(Qt.AlignmentFlag.AlignCenter)
        root.addWidget(brand)

        self.license_panel = QFrame()
        self.license_panel.setObjectName("LicenseHomePanel")
        self.license_layout = QHBoxLayout(self.license_panel)
        self.license_layout.setContentsMargins(18,12,18,12)
        self.license_layout.setSpacing(10)

        license_text = QVBoxLayout()
        license_text.setSpacing(1)
        self.license_heading = QLabel("LICENSE & ACCOUNT")
        self.license_heading.setObjectName("LicenseStatus")
        license_text.addWidget(self.license_heading)
        self.license_note = QLabel("Trial mode — 12 days remaining")
        self.license_note.setObjectName("Muted")
        license_text.addWidget(self.license_note)
        self.license_layout.addLayout(license_text,1)

        self.buy_btn = QPushButton("Buy License")
        self.buy_btn.setObjectName("LicenseBuy")
        self.buy_btn.clicked.connect(app.open_payment_preview)
        self.license_layout.addWidget(self.buy_btn)

        self.activate_btn = QPushButton("Activate License")
        self.activate_btn.setObjectName("LicenseAction")
        self.activate_btn.clicked.connect(app.open_activation_preview)
        self.license_layout.addWidget(self.activate_btn)

        self.recover_btn = QPushButton("Recover Existing Account")
        self.recover_btn.setObjectName("LicenseAction")
        self.recover_btn.clicked.connect(app.open_recovery_preview)
        self.license_layout.addWidget(self.recover_btn)

        self.manage_btn = QPushButton("Manage License")
        self.manage_btn.setObjectName("LicenseAction")
        self.manage_btn.clicked.connect(app.manage_license)
        self.manage_btn.hide()
        self.license_layout.addWidget(self.manage_btn)

        root.addWidget(self.license_panel)

        workflow_subtitle = QLabel("Data Workflow & Excel Automation")
        workflow_subtitle.setObjectName("HomeSubtitle")
        workflow_subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        root.addWidget(workflow_subtitle)

        welcome = QLabel("Choose a workflow to get started.")
        welcome.setObjectName("HomeWelcome")
        welcome.setAlignment(Qt.AlignmentFlag.AlignCenter)
        root.addWidget(welcome)

        self.cards = QGridLayout()
        self.cards.setHorizontalSpacing(14)
        self.cards.setVerticalSpacing(14)

        for module in modules:
            p = MODULE_PRESENTATION.get(
                module.info.module_id,
                (module.info.title.upper(),module.info.subtitle,module.info.description,"brand_mark.png","Primary","#173D2A")
            )
            card = ModuleCard(module.info.module_id,p[0],p[1],p[2],self.assets/p[3],p[4],p[5],app.open_module)
            self.module_cards.append(card)

        root.addLayout(self.cards)

        util = QFrame()
        util.setObjectName("UtilityPanel")
        ul = QGridLayout(util)
        ul.setContentsMargins(16,9,16,9)
        ul.setHorizontalSpacing(8)
        ul.setVerticalSpacing(5)

        heading = QLabel("▣  TOOLS & SUPPORT")
        heading.setObjectName("SectionTitle")
        ul.addWidget(heading,0,0,1,2)

        actions = [
            ("Check for Updates",app.check_for_updates),
            ("View Log Folder",app.open_log_folder),
        ]
        for index,(text,fn) in enumerate(actions):
            b = QPushButton(text)
            b.setObjectName("UtilityAction")
            b.clicked.connect(fn)
            ul.addWidget(b,1,index)

        root.addWidget(util)

        lower = QHBoxLayout()
        lower.setSpacing(10)

        recent = QGroupBox("RECENT PROJECTS")
        rl = QVBoxLayout(recent)
        recent_text = QLabel("Completed module runs will appear here as\nthe persistent project history is enabled.")
        recent_text.setWordWrap(True)
        rl.addWidget(recent_text)

        tips = QGroupBox("TIPS & GUIDANCE")
        tl = QVBoxLayout(tips)
        self._tips = [
            "New here? Start with a test workbook and review the Analysis summary before running a module.",
            "Split Reports can use the Current Excel Workbook View when you only want currently visible rows.",
            "Use Analysis before processing to confirm the worksheet, selected column and expected output count.",
            "Keep source workbooks unchanged until you have reviewed the generated output files.",
            "License recovery is intended for permitted device changes or reinstalls; keep your Recovery Code secure.",
        ]
        self._tip_index = 0
        self.tip_text = QLabel(self._tips[self._tip_index])
        self.tip_text.setWordWrap(True)
        self.tip_text.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        tl.addWidget(self.tip_text, 1)
        next_tip = QPushButton("Next Tip  ›")
        next_tip.setObjectName("UtilityAction")
        next_tip.clicked.connect(self._show_next_tip)
        tl.addWidget(next_tip, alignment=Qt.AlignmentFlag.AlignRight)

        status = QGroupBox("SYSTEM STATUS")
        sl = QVBoxLayout(status)
        self.system_license_status = QLabel("●  TRIAL MODE")
        self.system_license_status.setStyleSheet("color:#B06A10;font-size:14px;font-weight:800;")
        sl.addWidget(self.system_license_status)
        self.system_license_detail = QLabel("12 days remaining")
        self.system_license_detail.setStyleSheet("color:#24312B;font-size:13px;font-weight:700;")
        sl.addWidget(self.system_license_detail)
        self.system_license_note = QLabel("12-day trial period is active on this device.")
        self.system_license_note.setWordWrap(True)
        self.system_license_note.setObjectName("Muted")
        sl.addWidget(self.system_license_note)
        sl.addSpacing(4)
        ok = QLabel("Application systems operational")
        ok.setObjectName("Success")
        sl.addWidget(ok)
        sl.addWidget(QLabel(f"Version {VERSION}"))
        sl.addStretch()

        lower.addWidget(recent,2)
        lower.addWidget(tips,1)
        lower.addWidget(status,1)
        root.addLayout(lower)

        copyright_lbl = QLabel("© 2026 Golden Fir Solutions. All rights reserved.")
        copyright_lbl.setObjectName("Muted")
        copyright_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        root.addWidget(copyright_lbl)

        self._update_card_grid(4)
        self.set_license_state(getattr(app, "license_state", {"status":"trial","days_remaining":12}))

    def _show_next_tip(self):
        self._tip_index = (self._tip_index + 1) % len(self._tips)
        self.tip_text.setText(self._tips[self._tip_index])

    def set_license_state(self, state):
        status = state.get("status", "trial")
        active = status == "active"
        expired = status == "expired"
        unavailable = status in {"refunded", "revoked"}
        if active:
            license_type = state.get("license_type") or "Personal License"
            licensed_email = (state.get("email") or state.get("purchase_email") or "").strip()
            if licensed_email:
                self.license_note.setText(f"ACTIVE — {license_type} · Licensed to: {licensed_email}")
            else:
                self.license_note.setText(f"ACTIVE — {license_type}")
            self.buy_btn.hide()
            self.activate_btn.hide()
            self.recover_btn.hide()
            self.manage_btn.show()
            self.system_license_status.setText("●  LICENSE ACTIVE")
            self.system_license_status.setStyleSheet("color:#21833A;font-size:14px;font-weight:800;")
            self.system_license_detail.setText(license_type)
            self.system_license_note.setText("This device is activated in the current licensing state.")
        elif unavailable:
            label = "REFUNDED" if status == "refunded" else "REVOKED"
            self.license_note.setText(f"LICENSE {label} — access disabled")
            self.buy_btn.show()
            self.activate_btn.hide()
            self.recover_btn.hide()
            self.manage_btn.show()
            self.system_license_status.setText(f"●  LICENSE {label}")
            self.system_license_status.setStyleSheet("color:#B33A2B;font-size:14px;font-weight:800;")
            self.system_license_detail.setText("License access disabled")
            self.system_license_note.setText("Existing reports remain on this computer. A refunded or revoked license cannot be recovered.")
        elif expired:
            self.license_note.setText("TRIAL EXPIRED — license required")
            self.buy_btn.show()
            self.activate_btn.show()
            self.recover_btn.show()
            self.manage_btn.hide()
            self.system_license_status.setText("●  TRIAL EXPIRED")
            self.system_license_status.setStyleSheet("color:#B33A2B;font-size:14px;font-weight:800;")
            self.system_license_detail.setText("License required")
            self.system_license_note.setText("Activate a license to continue using the report modules.")
        else:
            days = int(state.get("days_remaining", 12))
            self.license_note.setText(f"Trial mode — {days} days remaining")
            self.buy_btn.show()
            self.activate_btn.show()
            self.recover_btn.show()
            self.manage_btn.hide()
            self.system_license_status.setText("●  TRIAL MODE")
            self.system_license_status.setStyleSheet("color:#B06A10;font-size:14px;font-weight:800;")
            self.system_license_detail.setText(f"{days} days remaining")
            if state.get("clock_rollback_detected"):
                self.system_license_note.setText("Trial time is based on the last verified local run time.")
            else:
                self.system_license_note.setText("12-day trial period is active on this device.")

    def _update_card_grid(self, columns):
        columns = max(1,min(int(columns),4))
        if columns == self._card_columns:
            return

        while self.cards.count():
            item = self.cards.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.setParent(self.content)

        for index,card in enumerate(self.module_cards):
            row = index // columns
            col = index % columns
            self.cards.addWidget(card,row,col)

        for col in range(4):
            self.cards.setColumnStretch(col,1 if col < columns else 0)

        self._card_columns = columns

    def resizeEvent(self, event):
        super().resizeEvent(event)
        viewport_width = self.scroll.viewport().width() if hasattr(self,"scroll") else self.width()

        if viewport_width >= 1080:
            columns = 4
        elif viewport_width >= 620:
            columns = 2
        else:
            columns = 1

        self._update_card_grid(columns)
