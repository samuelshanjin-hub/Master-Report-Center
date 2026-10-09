from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame, QPushButton, QApplication
from PyQt6.QtCore import Qt
from version import VERSION


class SimplePage(QWidget):
    def __init__(self, title, subtitle, body=""):
        super().__init__()
        root = QVBoxLayout(self)
        root.setContentsMargins(32, 28, 32, 28)
        root.setSpacing(10)
        t = QLabel(title)
        t.setObjectName("PageTitle")
        root.addWidget(t)
        s = QLabel(subtitle)
        s.setObjectName("PageSubtitle")
        s.setWordWrap(True)
        root.addWidget(s)
        panel = QFrame()
        panel.setObjectName("InfoPanel")
        pl = QVBoxLayout(panel)
        text = QLabel(body)
        text.setWordWrap(True)
        text.setAlignment(Qt.AlignmentFlag.AlignTop)
        pl.addWidget(text)
        root.addWidget(panel, 1)


class SettingsPage(QWidget):
    def __init__(self, app):
        super().__init__()
        self.app = app
        root = QVBoxLayout(self)
        root.setContentsMargins(32, 28, 32, 28)
        root.setSpacing(12)

        title = QLabel("Settings")
        title.setObjectName("PageTitle")
        root.addWidget(title)

        subtitle = QLabel("Application, configuration and diagnostic preferences.")
        subtitle.setObjectName("PageSubtitle")
        subtitle.setWordWrap(True)
        root.addWidget(subtitle)

        config_panel = QFrame()
        config_panel.setObjectName("InfoPanel")
        cp = QVBoxLayout(config_panel)
        heading = QLabel("Configuration backup & transfer")
        heading.setStyleSheet("font-size:15px;font-weight:800;")
        cp.addWidget(heading)

        info = QLabel(
            "Export Configuration creates a JSON backup of supported application settings for safekeeping or transfer to another installation.\n\n"
            "Import Configuration loads a previously exported configuration file. Importing may change current application settings, so review the file source and keep a backup before continuing."
        )
        info.setWordWrap(True)
        cp.addWidget(info)

        warning = QLabel(
            "Important: You are responsible for reviewing configuration files before importing them. Incorrect, incompatible or modified configuration files may change application behaviour. "
            "This does not limit any rights or remedies that cannot lawfully be excluded, including rights under the Australian Consumer Law."
        )
        warning.setWordWrap(True)
        warning.setObjectName("Muted")
        cp.addWidget(warning)

        actions = QHBoxLayout()
        import_btn = QPushButton("Import Configuration")
        import_btn.clicked.connect(app.import_configuration)
        actions.addWidget(import_btn)
        export_btn = QPushButton("Export Configuration")
        export_btn.clicked.connect(app.export_configuration)
        actions.addWidget(export_btn)
        actions.addStretch()
        cp.addLayout(actions)
        root.addWidget(config_panel)

        diagnostics = QFrame()
        diagnostics.setObjectName("InfoPanel")
        dl = QVBoxLayout(diagnostics)
        dh = QLabel("Diagnostics")
        dh.setStyleSheet("font-size:15px;font-weight:800;")
        dl.addWidget(dh)
        dt = QLabel("Use the log folder when troubleshooting or when support asks for diagnostic information.")
        dt.setWordWrap(True)
        dl.addWidget(dt)
        log_btn = QPushButton("View Log Folder")
        log_btn.clicked.connect(app.open_log_folder)
        dl.addWidget(log_btn, alignment=Qt.AlignmentFlag.AlignLeft)
        root.addWidget(diagnostics)
        root.addStretch()


class LicensePage(QWidget):
    def __init__(self, app):
        super().__init__()
        self.app = app
        root = QVBoxLayout(self)
        root.setContentsMargins(32, 28, 32, 28)
        root.setSpacing(12)

        title = QLabel("License")
        title.setObjectName("PageTitle")
        root.addWidget(title)

        subtitle = QLabel("Manage your Master Report Center All In One license.")
        subtitle.setObjectName("PageSubtitle")
        root.addWidget(subtitle)

        self.status_panel = QFrame()
        self.status_panel.setObjectName("InfoPanel")
        sl = QVBoxLayout(self.status_panel)
        self.status_title = QLabel("TRIAL MODE")
        self.status_title.setStyleSheet("font-size:15px;font-weight:800;color:#B06A10;")
        sl.addWidget(self.status_title)
        self.status_detail = QLabel("12 days remaining")
        self.status_detail.setStyleSheet("font-size:13px;font-weight:700;")
        sl.addWidget(self.status_detail)
        self.status_note = QLabel("Your 12-day trial is active on this device.")
        self.status_note.setObjectName("Muted")
        self.status_note.setWordWrap(True)
        sl.addWidget(self.status_note)
        root.addWidget(self.status_panel)

        self.account_panel = QFrame()
        self.account_panel.setObjectName("InfoPanel")
        apl = QVBoxLayout(self.account_panel)
        account_heading = QLabel("License & recovery details")
        account_heading.setStyleSheet("font-size:15px;font-weight:800;")
        apl.addWidget(account_heading)

        self.account_email = QLabel()
        self.account_email.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.account_license_id = QLabel()
        self.account_license_id.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.account_recovery = QLabel()
        self.account_recovery.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.account_recovery.setStyleSheet("font-size:14px;font-weight:800;color:#173F32;")
        self.account_migrations = QLabel()
        self.account_migrations.setObjectName("Muted")
        self.account_migrations.setWordWrap(True)
        apl.addWidget(self.account_email)
        apl.addWidget(self.account_license_id)
        apl.addWidget(self.account_recovery)
        apl.addWidget(self.account_migrations)

        copy_row = QHBoxLayout()
        self.copy_email_btn = QPushButton("Copy Purchase Email")
        self.copy_email_btn.clicked.connect(lambda: QApplication.clipboard().setText((getattr(self, '_active_email', '') or '')))
        copy_row.addWidget(self.copy_email_btn)
        self.copy_code_btn = QPushButton("Copy Recovery Code")
        self.copy_code_btn.clicked.connect(lambda: QApplication.clipboard().setText((getattr(self, '_active_recovery_code', '') or '')))
        copy_row.addWidget(self.copy_code_btn)
        copy_row.addStretch()
        apl.addLayout(copy_row)
        root.addWidget(self.account_panel)

        plans = QFrame()
        plans.setObjectName("InfoPanel")
        pl = QVBoxLayout(plans)
        heading = QLabel("Purchase & account options")
        heading.setStyleSheet("font-size:15px;font-weight:800;")
        pl.addWidget(heading)
        details = QLabel(
            "Personal License â€” AUD 49 one-time\n"
            "For individual professional use. Up to 4 device migrations or rebindings per 12 months.\n\n"
            "Company License â€” AUD 149 one-time\n"
            "For business-owned use. No annual device recovery or rebinding count limit.\n\n"
            "One-time purchase. No subscription. Australian Consumer Law rights remain applicable."
        )
        details.setWordWrap(True)
        pl.addWidget(details)

        actions = QHBoxLayout()
        self.buy_btn = QPushButton("Buy License")
        self.buy_btn.clicked.connect(app.open_payment_preview)
        actions.addWidget(self.buy_btn)
        self.activate_btn = QPushButton("Activate License")
        self.activate_btn.clicked.connect(app.open_activation_preview)
        actions.addWidget(self.activate_btn)
        self.recover_btn = QPushButton("Recover Existing Account")
        self.recover_btn.clicked.connect(app.open_recovery_preview)
        actions.addWidget(self.recover_btn)
        actions.addStretch()
        pl.addLayout(actions)
        root.addWidget(plans)

        self.refund_panel = QFrame()
        self.refund_panel.setObjectName("InfoPanel")
        rfl = QVBoxLayout(self.refund_panel)
        refund_heading = QLabel("Refund request")
        refund_heading.setStyleSheet("font-size:15px;font-weight:800;")
        rfl.addWidget(refund_heading)
        refund = QLabel(
            "Paid customers can submit a refund request using the purchase email, Activation / Recovery Code and receipt or purchase ID. "
            "Submitting a request does not immediately revoke the license. The request is reviewed first and the license is revoked only after an approved refund is confirmed by the payment provider."
        )
        refund.setWordWrap(True)
        refund.setObjectName("Muted")
        rfl.addWidget(refund)
        refund_button = QPushButton("Request a Refund")
        refund_button.clicked.connect(app.open_refund_preview)
        rfl.addWidget(refund_button, alignment=Qt.AlignmentFlag.AlignLeft)
        root.addWidget(self.refund_panel)
        root.addStretch()

        self.set_license_state(getattr(app, "license_state", {"status":"trial","days_remaining":12}))

    def set_license_state(self, state):
        status = state.get("status", "trial")
        active = status == "active"
        expired = status == "expired"
        unavailable = status in {"refunded", "revoked"}
        if active:
            license_type = state.get("license_type") or "Personal License"
            licensed_email = (state.get("email") or state.get("purchase_email") or "").strip()
            self.status_title.setText("LICENSE ACTIVE")
            self.status_title.setStyleSheet("font-size:15px;font-weight:800;color:#21833A;")
            if licensed_email:
                self.status_detail.setText(f"{license_type} Â· Licensed to: {licensed_email}")
            else:
                self.status_detail.setText(license_type)
            self.status_note.setText("This device is activated. Keep the Recovery Code safe; it is required for account recovery or a permitted device transfer.")
            recovery_code = (state.get("recovery_code") or "").strip()
            recovery_last4 = (state.get("recovery_code_last4") or "").strip()
            license_id = (state.get("license_id") or "").strip()
            self._active_email = licensed_email
            self._active_recovery_code = recovery_code
            self.account_email.setText(f"Purchase Email: {licensed_email or 'Not available'}")
            self.account_license_id.setText(f"License ID: {license_id or 'Not available'}")
            if recovery_code:
                self.account_recovery.setText(f"Recovery Code: {recovery_code}")
                self.copy_code_btn.setEnabled(True)
            elif recovery_last4:
                self.account_recovery.setText(f"Recovery Code: â€¢â€¢â€¢â€¢{recovery_last4} (full code is not stored by the server)")
                self.copy_code_btn.setEnabled(False)
            else:
                self.account_recovery.setText("Recovery Code: Not available on this device")
                self.copy_code_btn.setEnabled(False)

            if "Company" in license_type:
                self.account_migrations.setText("Device migrations / rebindings: No annual limit (events are still recorded).")
            else:
                used = state.get("migrations_used_12_months")
                remaining = state.get("migrations_remaining_12_months")
                if used is not None and remaining is not None:
                    self.account_migrations.setText(
                        f"Device migrations / rebindings: {remaining} of 4 remaining in the current rolling 12-month period "
                        f"({used} used)."
                    )
                else:
                    self.account_migrations.setText(
                        "Device migrations / rebindings: up to 4 per rolling 12 months. "
                        "The exact remaining count will refresh after the next verified activation or recovery."
                    )
            self.account_panel.show()
            self.buy_btn.hide()
            self.activate_btn.hide()
            self.recover_btn.hide()
            self.refund_panel.show()
        elif unavailable:
            label = "REFUNDED" if status == "refunded" else "REVOKED"
            self.status_title.setText(f"LICENSE {label}")
            self.status_title.setStyleSheet("font-size:15px;font-weight:800;color:#B33A2B;")
            self.status_detail.setText("License access disabled")
            self.status_note.setText("This paid entitlement is no longer active. Existing reports and files remain on this computer, but the report modules are locked.")
            licensed_email = (state.get("email") or state.get("purchase_email") or "").strip()
            license_id = (state.get("license_id") or "").strip()
            self._active_email = licensed_email
            self._active_recovery_code = ""
            self.account_email.setText(f"Purchase Email: {licensed_email or 'Not available'}")
            self.account_license_id.setText(f"License ID: {license_id or 'Not available'}")
            self.account_recovery.setText("Recovery Code: disabled for this license status")
            self.account_migrations.setText("Device activation and recovery are disabled for a refunded or revoked license.")
            self.copy_code_btn.setEnabled(False)
            self.account_panel.show()
            self.buy_btn.show()
            self.activate_btn.hide()
            self.recover_btn.hide()
            self.refund_panel.hide()
        elif expired:
            self.status_title.setText("TRIAL EXPIRED")
            self.status_title.setStyleSheet("font-size:15px;font-weight:800;color:#B33A2B;")
            self.status_detail.setText("License required")
            self.status_note.setText("The 12-day trial has ended. Buy a license, activate an existing license, or recover your account to continue using the report modules.")
            self.account_panel.hide()
            self.buy_btn.show()
            self.activate_btn.show()
            self.recover_btn.show()
            self.refund_panel.hide()
        else:
            days = int(state.get("days_remaining", 12))
            self.status_title.setText("TRIAL MODE")
            self.status_title.setStyleSheet("font-size:15px;font-weight:800;color:#B06A10;")
            self.status_detail.setText(f"{days} days remaining")
            if state.get("clock_rollback_detected"):
                self.status_note.setText("Trial time is based on the last verified local run time.")
            else:
                self.status_note.setText("Your 12-day trial is active on this device.")
            self.account_panel.hide()
            self.buy_btn.show()
            self.activate_btn.show()
            self.recover_btn.show()
            self.refund_panel.hide()


class AboutPage(QWidget):
    def __init__(self):
        super().__init__()
        root = QVBoxLayout(self)
        root.setContentsMargins(32, 28, 32, 28)
        root.setSpacing(10)

        title = QLabel("About")
        title.setObjectName("PageTitle")
        root.addWidget(title)

        subtitle = QLabel("Golden Fir Report Center")
        subtitle.setObjectName("PageSubtitle")
        root.addWidget(subtitle)

        panel = QFrame()
        panel.setObjectName("InfoPanel")
        pl = QVBoxLayout(panel)
        body = QLabel(
            f"Master Report Center All In One\n"
            f"Version {VERSION}\n\n"
            "Golden Fir Solutions\n"
            "Automate â€¢ Simplify â€¢ Empower\n\n"
            "Professional Excel reporting and data workflow tools.\n"
            "Split Reports â€¢ Consolidate Reports â€¢ Synchronize Reports â€¢ Lookup Data"
        )
        body.setWordWrap(True)
        body.setAlignment(Qt.AlignmentFlag.AlignTop)
        pl.addWidget(body)
        root.addWidget(panel, 1)

