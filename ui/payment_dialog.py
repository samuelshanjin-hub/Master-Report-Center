from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

from PyQt6.QtCore import Qt, QUrl, QTimer
from PyQt6.QtGui import QDesktopServices, QGuiApplication, QPixmap
from version import VERSION
from core.services.backend_client import BackendClientError, LicensingBackendClient

from PyQt6.QtWidgets import (
    QApplication, QButtonGroup, QCheckBox, QDialog, QFrame, QGridLayout, QHBoxLayout,
    QLabel, QLineEdit, QMessageBox, QPushButton, QRadioButton, QTextBrowser,
    QVBoxLayout, QWidget, QComboBox, QTextEdit, QScrollArea, QSizePolicy
)


BRAND_GREEN = "#1F5E46"
BRAND_GREEN_DARK = "#173F32"
BRAND_GREEN_SOFT = "#EEF5F1"
CHARCOAL = "#24312B"
PAGE_BG = "#F3F6F4"
BORDER = "#D6E0DA"
GOLD = "#B58A2A"
TERMS_VERSION = "MRC-PURCHASE-1.0"


class PurchaseTermsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Purchase Terms")
        self.resize(760, 620)
        self.setModal(True)
        self.setStyleSheet(f"""
            QDialog {{ background:{PAGE_BG}; }}
            QLabel#Title {{ color:{BRAND_GREEN_DARK}; font-size:22px; font-weight:800; }}
            QTextBrowser {{ background:#FFFFFF; border:1px solid {BORDER}; border-radius:9px; padding:10px; color:{CHARCOAL}; }}
            QPushButton {{ min-height:36px; border-radius:8px; padding:6px 14px; }}
            QPushButton#Primary {{ background:{BRAND_GREEN}; color:#FFFFFF; border:none; font-weight:700; }}
        """)

        root = QVBoxLayout(self)
        root.setContentsMargins(24, 22, 24, 22)
        root.setSpacing(12)

        title = QLabel("Master Report Center All In One — Purchase Terms")
        title.setObjectName("Title")
        root.addWidget(title)

        terms = QTextBrowser()
        terms.setOpenExternalLinks(False)
        terms.setHtml(f"""
        <p><b>Terms version:</b> {TERMS_VERSION}</p>
        <p><b>Product:</b> Master Report Center All In One</p>
        <p><b>Purchase model:</b> One-time purchase. No subscription.</p>

        <h3>License options</h3>
        <p><b>Personal License — AUD 49</b><br>
        Intended for individual professional use. A Personal License may be migrated or rebound to a replacement or reinstalled device up to four times in each 12-month period. Normal reopening, software updates and re-use on the same recognised device do not consume a migration.</p>

        <p><b>Company License — AUD 149</b><br>
        Intended for business-owned use. The Company account does not have an annual device recovery / rebinding count limit. Final deployment and authorised-use scope will be stated in the live licence terms.</p>

        <h3>Activation and recovery</h3>
        <p>After confirmed payment, the purchased entitlement is associated with the purchase email and an Activation / Recovery Code is provided for permitted recovery and device changes. The code should be kept secure.</p>

        <h3>Payment and privacy</h3>
        <p>Payment will be processed by an established third-party payment provider. Golden Fir will not directly collect or store the purchaser's full payment-card number. Available payment methods can vary by device, browser, country and payment provider.</p>

        <h3>User configuration changes</h3>
        <p>Users are responsible for reviewing configuration files before importing them and for changes made through user-selected configuration settings. Incorrect, incompatible or modified configuration files may change application behaviour. This responsibility does not exclude, restrict or modify any rights or remedies that cannot lawfully be excluded.</p>

        <h3>Australian Consumer Law</h3>
        <p>Nothing in these terms excludes, restricts or modifies any consumer guarantee, right or remedy that cannot lawfully be excluded under the Australian Consumer Law. Remedies for faulty or non-compliant software will be handled in accordance with applicable Australian consumer law. Any separate change-of-mind policy will not limit statutory consumer rights.</p>

        <h3>Purchase record</h3>
        <p>The licensing system records the selected licence, price, purchase email, payment transaction reference, accepted terms version and acceptance time for the transaction.</p>

        <h3>Seller details</h3>
        <p>Support and refund information is available in the application and through Golden Fir Solutions support. Statutory consumer rights remain unaffected.</p>
        """)
        root.addWidget(terms, 1)

        row = QHBoxLayout()
        row.addStretch()
        close = QPushButton("Close")
        close.setObjectName("Primary")
        close.clicked.connect(self.accept)
        row.addWidget(close)
        root.addLayout(row)


class PaymentDialog(QDialog):
    def __init__(self, assets_dir: Path, parent=None):
        super().__init__(parent)
        self.assets_dir = Path(assets_dir)
        self.setWindowTitle("Golden Fir Secure Checkout")
        self.setModal(True)
        self.selected_license = "Personal License"
        self.selected_price = "AUD 49.00"
        self.backend_client = LicensingBackendClient()
        self.last_purchase_intent_id = None
        self.payment_timer = QTimer(self)
        self.payment_timer.setInterval(2000)
        self.payment_timer.timeout.connect(self._check_payment_and_activate)
        self._activation_in_progress = False

        # Responsive checkout sizing.  The previous fixed 920x680 minimum could
        # push the top of the dialog off-screen on smaller laptop displays.
        screen = QGuiApplication.screenAt(self.parentWidget().frameGeometry().center()) if self.parentWidget() else None
        if screen is None:
            screen = QGuiApplication.primaryScreen()
        available = screen.availableGeometry() if screen else None
        screen_w = available.width() if available else 1366
        screen_h = available.height() if available else 768

        self._compact_checkout = screen_h < 800
        button_height = 34 if self._compact_checkout else 40
        line_padding = 7 if self._compact_checkout else 9
        target_w = min(1080, max(800, screen_w - 70))
        target_h = min(760, max(540, screen_h - 70))
        self.setMinimumSize(780, 520)
        self.resize(target_w, target_h)

        self.setStyleSheet(f"""
            QDialog {{ background:{PAGE_BG}; }}
            QFrame#CheckoutCard {{ background:#FFFFFF; border:1px solid {BORDER}; border-radius:14px; }}
            QFrame#InfoPanel {{ background:#F8FAF9; border:1px solid #DFE7E2; border-radius:12px; }}
            QFrame#LicenseCard {{ background:#FFFFFF; border:1px solid #D4DED8; border-radius:10px; }}
            QFrame#LicenseCardSelected {{ background:{BRAND_GREEN_SOFT}; border:2px solid {BRAND_GREEN}; border-radius:10px; }}
            QLabel#Brand {{ color:{BRAND_GREEN_DARK}; font-size:13px; font-weight:800; letter-spacing:1px; }}
            QLabel#WhiteTitle {{ color:{CHARCOAL}; font-size:26px; font-weight:800; }}
            QLabel#WhiteText {{ color:#53625B; font-size:13px; }}
            QLabel#Title {{ color:{CHARCOAL}; font-size:23px; font-weight:800; }}
            QLabel#Sub {{ color:#66756E; font-size:13px; }}
            QLabel#Trust {{ color:#52615A; font-size:12px; }}
            QLabel#Badge {{ background:#FFF8E7; color:#7A5A12; border:1px solid #E8D6A4; border-radius:9px; padding:5px 9px; font-weight:700; }}
            QLabel#Wallet {{ background:#F7F9F8; color:#33443C; border:1px solid #DCE5E0; border-radius:8px; padding:8px 10px; font-weight:700; }}
            QPushButton {{ min-height:{button_height}px; border-radius:8px; padding:7px 14px; font-weight:700; }}
            QPushButton#PayPal {{ background:#FFC439; color:#142C8E; border:none; }}
            QPushButton#Card {{ background:{BRAND_GREEN}; color:#FFFFFF; border:none; }}
            QPushButton#Ghost {{ background:#FFFFFF; color:{BRAND_GREEN_DARK}; border:1px solid #BFCFC6; }}
            QPushButton#Ghost:disabled {{ color:#99A59F; background:#F4F6F5; border-color:#DDE4E0; }}
            QPushButton#PayPal:disabled, QPushButton#Card:disabled {{ background:#D9DFDC; color:#8A9690; }}
            QPushButton#Link {{ background:transparent; color:{BRAND_GREEN}; border:none; text-decoration:underline; }}
            QLineEdit {{ background:#FFFFFF; border:1px solid #C9D5CE; border-radius:7px; padding:{line_padding}px; min-height:20px; }}
            QRadioButton {{ color:{CHARCOAL}; font-weight:700; spacing:8px; }}
            QCheckBox {{ color:{CHARCOAL}; spacing:8px; }}
        """)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(18 if self._compact_checkout else 30, 14 if self._compact_checkout else 24, 18 if self._compact_checkout else 30, 14 if self._compact_checkout else 24)
        outer.setSpacing(8 if self._compact_checkout else 12)

        top = QHBoxLayout()
        brand = QLabel("GOLDEN FIR  •  SECURE CHECKOUT")
        brand.setObjectName("Brand")
        top.addWidget(brand)
        top.addStretch()
        demo = QLabel("SECURE PAYPAL CHECKOUT")
        demo.setObjectName("Badge")
        top.addWidget(demo)
        outer.addLayout(top)

        card = QFrame(); card.setObjectName("CheckoutCard")
        card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        card_layout = QHBoxLayout(card)
        card_layout.setContentsMargins(12 if self._compact_checkout else 18, 12 if self._compact_checkout else 18, 12 if self._compact_checkout else 18, 12 if self._compact_checkout else 18)
        card_layout.setSpacing(12 if self._compact_checkout else 18)

        left = QFrame(); left.setObjectName("InfoPanel"); left.setMinimumWidth(255 if self._compact_checkout else 300)
        ll = QVBoxLayout(left)
        ll.setContentsMargins(14 if self._compact_checkout else 22, 14 if self._compact_checkout else 22, 14 if self._compact_checkout else 22, 14 if self._compact_checkout else 22)
        ll.setSpacing(7 if self._compact_checkout else 11)

        logo = QLabel()
        logo_path = self.assets_dir / "golden_fir_wordmark.png"
        if logo_path.exists():
            pm = QPixmap(str(logo_path))
            logo.setPixmap(pm.scaled(190 if self._compact_checkout else 235, 68 if self._compact_checkout else 90, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        ll.addWidget(logo)

        prod = QLabel("Master Report Center\nAll In One")
        prod.setObjectName("WhiteTitle"); prod.setWordWrap(True)
        ll.addWidget(prod)

        copy = QLabel("One-time software purchase. No subscription. Choose the licence that fits how the software will be owned and managed.")
        copy.setObjectName("WhiteText"); copy.setWordWrap(True)
        ll.addWidget(copy)

        for text in (
            "✓ Payment handled by established providers",
            "✓ Golden Fir will not store full card numbers",
            "✓ Desktop or mobile secure checkout",
            "✓ Receipt + Activation / Recovery Code after payment",
            "✓ Australian Consumer Law rights remain applicable",
        ):
            item = QLabel(text); item.setObjectName("WhiteText"); item.setWordWrap(True); ll.addWidget(item)
        ll.addStretch()
        card_layout.addWidget(left, 3)

        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(5 if self._compact_checkout else 8, 2 if self._compact_checkout else 4, 5 if self._compact_checkout else 8, 2 if self._compact_checkout else 4)
        rl.setSpacing(6 if self._compact_checkout else 9)

        title = QLabel("Choose your licence")
        title.setObjectName("Title")
        rl.addWidget(title)

        self.license_group = QButtonGroup(self)
        self.license_group.setExclusive(True)
        license_row = QHBoxLayout(); license_row.setSpacing(10)
        self.personal_card, self.personal_radio = self._license_card(
            "Personal License", "AUD 49", "Individual professional use", "Up to 4 device migrations / rebindings per 12 months."
        )
        self.company_card, self.company_radio = self._license_card(
            "Company License", "AUD 149", "Business-owned use", "No annual device recovery / rebinding count limit."
        )
        self.license_group.addButton(self.personal_radio, 0)
        self.license_group.addButton(self.company_radio, 1)
        self.personal_radio.setChecked(True)
        self.personal_radio.toggled.connect(self._update_license_selection)
        self.company_radio.toggled.connect(self._update_license_selection)
        license_row.addWidget(self.personal_card, 1)
        license_row.addWidget(self.company_card, 1)
        rl.addLayout(license_row)
        self._update_license_selection()

        email = QLineEdit()
        email.setPlaceholderText("Purchase email — used for receipt and license recovery")
        self.email_input = email
        self.email_input.textChanged.connect(self._update_payment_enabled)
        rl.addWidget(email)

        terms_row = QHBoxLayout()
        self.terms_check = QCheckBox("I have read and agree to the Purchase Terms and License Conditions.")
        self.terms_check.toggled.connect(self._update_payment_enabled)
        terms_row.addWidget(self.terms_check, 1)
        view_terms = QPushButton("View Terms")
        view_terms.setObjectName("Link")
        view_terms.clicked.connect(lambda: PurchaseTermsDialog(self).exec())
        terms_row.addWidget(view_terms)
        rl.addLayout(terms_row)

        payment_title = QLabel("Secure payment routes")
        payment_title.setStyleSheet("font-size:15px;font-weight:800;color:#24312B;")
        rl.addWidget(payment_title)

        self.checkout_btn = QPushButton("Open Secure Checkout")
        self.checkout_btn.setObjectName("PayPal")
        self.checkout_btn.clicked.connect(self._open_live_checkout)
        rl.addWidget(self.checkout_btn)

        wallets = QLabel("PayPal-hosted checkout may offer:   PayPal   •   Pay Later   •   Apple Pay   •   Debit / Credit Card")
        wallets.setObjectName("Wallet")
        wallets.setWordWrap(True)
        rl.addWidget(wallets)
        wallet_note = QLabel("Available payment methods are provided by PayPal and may vary by device, browser, account eligibility and region. Payment is completed on PayPal's secure checkout page, not inside this desktop application.")
        wallet_note.setObjectName("Trust"); wallet_note.setWordWrap(True)
        rl.addWidget(wallet_note)

        trust = QLabel(
            "Checkout remains disabled until a purchase email is entered and the Purchase Terms are accepted. "
            "A unique PayPal checkout is created securely for each purchase through the Golden Fir licensing service."
        )
        trust.setObjectName("Trust"); trust.setWordWrap(True)
        rl.addWidget(trust)
        rl.addStretch()

        card_layout.addWidget(right, 7)

        self.checkout_scroll = QScrollArea()
        self.checkout_scroll.setWidgetResizable(True)
        self.checkout_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.checkout_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.checkout_scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")
        self.checkout_scroll.setWidget(card)
        outer.addWidget(self.checkout_scroll, 1)

        bottom = QHBoxLayout(); bottom.addStretch()
        close = QPushButton("Close"); close.setObjectName("Ghost"); close.clicked.connect(self.reject)
        bottom.addWidget(close); outer.addLayout(bottom)

        self._update_payment_enabled()

    def showEvent(self, event):
        super().showEvent(event)
        # Always start at the top of checkout content on small displays.
        if hasattr(self, "checkout_scroll"):
            self.checkout_scroll.verticalScrollBar().setValue(0)

        screen = self.screen() or QGuiApplication.primaryScreen()
        if screen:
            available = screen.availableGeometry()
            frame = self.frameGeometry()
            frame.moveCenter(available.center())
            self.move(frame.topLeft())

    def _license_card(self, name: str, price: str, subtitle: str, detail: str):
        frame = QFrame(); frame.setObjectName("LicenseCard")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(10 if self._compact_checkout else 14, 8 if self._compact_checkout else 12, 10 if self._compact_checkout else 14, 8 if self._compact_checkout else 12)
        layout.setSpacing(3 if self._compact_checkout else 5)
        radio = QRadioButton(name)
        layout.addWidget(radio)
        p = QLabel(price + "  •  one-time")
        p.setStyleSheet(f"color:{BRAND_GREEN_DARK};font-size:{17 if self._compact_checkout else 19}px;font-weight:800;")
        layout.addWidget(p)
        s = QLabel(subtitle); s.setStyleSheet("color:#3C4B44;font-weight:700;"); s.setWordWrap(True)
        layout.addWidget(s)
        d = QLabel(detail); d.setObjectName("Trust"); d.setWordWrap(True)
        layout.addWidget(d)
        layout.addStretch()
        return frame, radio

    def _update_license_selection(self):
        personal = self.personal_radio.isChecked()
        self.personal_card.setObjectName("LicenseCardSelected" if personal else "LicenseCard")
        self.company_card.setObjectName("LicenseCardSelected" if not personal else "LicenseCard")
        for card in (self.personal_card, self.company_card):
            card.style().unpolish(card); card.style().polish(card)
        if personal:
            self.selected_license = "Personal License"
            self.selected_price = "AUD 49.00"
        else:
            self.selected_license = "Company License"
            self.selected_price = "AUD 149.00"
    def _update_payment_enabled(self):
        enabled = self.terms_check.isChecked() and bool(self.email_input.text().strip())
        self.checkout_btn.setEnabled(enabled)

    @staticmethod
    def _looks_like_email(value: str) -> bool:
        value = value.strip()
        return bool(re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", value))

    def _product_code(self) -> str:
        return "MRC-PERSONAL" if self.selected_license == "Personal License" else "MRC-COMPANY"

    def _create_purchase_intent(self, email: str) -> dict | None:
        try:
            result = self.backend_client.create_purchase_intent(
                email=email,
                product_code=self._product_code(),
                terms_version=TERMS_VERSION,
                app_version=VERSION,
            )
        except BackendClientError as exc:
            QMessageBox.critical(
                self,
                "Checkout Not Started",
                "A Purchase Intent could not be created, so PayPal has NOT been opened.\n\n"
                f"{exc}\n\n"
                "Please check your internet connection and try again. If the problem continues, contact Golden Fir Solutions support.",
            )
            return None

        self.last_purchase_intent_id = result.get("purchase_intent_id")
        return result

    def _create_paypal_order(self, purchase_intent_id: str) -> dict | None:
        try:
            return self.backend_client.create_paypal_order(purchase_intent_id)
        except BackendClientError as exc:
            QMessageBox.critical(
                self,
                "PayPal Order Not Created",
                "The Purchase Intent was created, but the licensing service could not create a PayPal Order.\n\n"
                f"Purchase Intent: {purchase_intent_id}\n\n"
                f"{exc}\n\n"
                "Please try again. If the problem continues, contact Golden Fir Solutions support.",
            )
            return None

    def _open_live_checkout(self):
        email = self.email_input.text().strip()
        if not email:
            QMessageBox.warning(self, "Purchase Email", "Enter the purchase email that should receive the receipt and license recovery information.")
            self.email_input.setFocus()
            return
        if not self._looks_like_email(email):
            QMessageBox.warning(self, "Purchase Email", "Enter a valid purchase email address before continuing.")
            self.email_input.setFocus()
            return
        if not self.terms_check.isChecked():
            QMessageBox.warning(self, "Purchase Terms", "Please review and accept the Purchase Terms before continuing.")
            return

        answer = QMessageBox.question(
            self,
            "Create PayPal Order",
            f"You are about to create a Golden Fir Purchase Intent and a unique PayPal Order for:\n\n"
            f"{self.selected_license} — {self.selected_price}\n"
            f"Purchase email: {email}\n"
            f"Terms version: {TERMS_VERSION}\n\n"
            "Continue to the PayPal checkout page?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return

        intent = self._create_purchase_intent(email)
        if not intent:
            return

        intent_id = str(intent.get("purchase_intent_id", ""))
        order = self._create_paypal_order(intent_id)
        if not order:
            return

        order_id = str(order.get("paypal_order_id", ""))
        approval_url = str(order.get("approval_url", ""))
        environment = str(order.get("environment", "production")).upper()

        opened = QDesktopServices.openUrl(QUrl(approval_url))
        if not opened:
            QMessageBox.warning(
                self,
                "Open PayPal Checkout",
                f"Purchase Intent {intent_id} and PayPal Order {order_id} were created, but the approval page could not be opened in your default browser.\n\n"
                "No license has been activated.",
            )
            return

        QMessageBox.information(
            self,
            "PayPal Order Created",
            f"Purchase Intent:\n{intent_id}\n\n"
            f"PayPal Order:\n{order_id}\n\n"
            f"Environment: {environment}\n\n"
            "Complete the PayPal checkout in the browser, then return to Master Report Center. "
            "This window will check the confirmed payment and activate this device automatically.",
        )
        self.payment_timer.start()

    def _check_payment_and_activate(self):
        if self._activation_in_progress or not self.last_purchase_intent_id:
            return
        try:
            status = self.backend_client.get_purchase_status(self.last_purchase_intent_id)
        except BackendClientError:
            return

        payment_status = str(status.get("payment_status", "")).upper()
        if payment_status != "PAID":
            return

        self._activation_in_progress = True
        self.payment_timer.stop()
        try:
            activation = self.backend_client.activate_paid_purchase(self.last_purchase_intent_id)
        except BackendClientError as exc:
            self._activation_in_progress = False
            QMessageBox.critical(
                self,
                "License Activation Failed",
                "The payment is confirmed, but this device could not be activated.\n\n"
                f"{exc}",
            )
            return

        app = self.parent()
        if hasattr(app, "apply_license_activation"):
            app.apply_license_activation(activation)

        recovery_code = str(activation.get("recovery_code", "") or "").strip()
        license_id = str(activation.get("license_id", "") or "").strip()
        email = str(activation.get("email", "") or "").strip()
        license_type = str(activation.get("license_type", "") or "").title()

        if recovery_code:
            message = (
                "Payment confirmed and this device is now activated.\n\n"
                f"License: {license_type}\n"
                f"Purchase Email: {email}\n"
                f"License ID: {license_id}\n\n"
                f"Recovery Code:\n{recovery_code}\n\n"
                "Keep this Recovery Code safe. The Golden Fir backend stores only a one-way hash of the code. "
                "You can copy the code later from the License page on this device."
            )
        else:
            message = (
                "Payment confirmed and this device is already activated for this license.\n\n"
                f"License: {license_type}\n"
                f"Purchase Email: {email}\n"
                f"License ID: {license_id}"
            )

        QMessageBox.information(self, "License Activated", message)
        self.accept()



class ActivationDialog(QDialog):
    def __init__(self, title: str, mode: str, parent=None):
        super().__init__(parent)
        self.mode = "activate" if mode == "activate" else "recover"
        self.setWindowTitle(title)
        self.resize(540, 360)
        self.setModal(True)
        self.setStyleSheet(f"""
            QDialog {{ background:#F6F8F7; }}
            QLabel#Title {{ color:{BRAND_GREEN_DARK}; font-size:22px; font-weight:800; }}
            QLabel#Sub {{ color:#5F6F67; }}
            QLineEdit {{ background:white; border:1px solid #CDD8D2; border-radius:7px; padding:9px; }}
            QPushButton#Primary {{ background:{BRAND_GREEN};color:white;border:none;border-radius:8px;padding:9px 14px;font-weight:700; }}
        """)
        root = QVBoxLayout(self); root.setContentsMargins(28, 26, 28, 26); root.setSpacing(10)
        t = QLabel(title); t.setObjectName("Title"); root.addWidget(t)
        if self.mode == "recover":
            body = (
                "Enter the original Purchase Email and Recovery Code to restore an existing paid license on this device. "
                "A new-device Personal recovery counts toward the 4 migrations/rebindings allowed in each 12-month period; same-device restoration does not count."
            )
        else:
            body = (
                "Enter the Purchase Email and Recovery Code issued for the paid license. "
                "The licensing service will verify the entitlement before activating this device."
            )
        s = QLabel(body); s.setObjectName("Sub"); s.setWordWrap(True); root.addWidget(s)

        self.email = QLineEdit(); self.email.setPlaceholderText("Purchase email")
        self.code = QLineEdit(); self.code.setPlaceholderText("Recovery Code (GFRC-....)")
        root.addWidget(self.email); root.addWidget(self.code)

        note = QLabel("The Recovery Code is verified by the Golden Fir licensing service. It is not stored in plaintext on the server.")
        note.setObjectName("Sub"); note.setWordWrap(True); root.addWidget(note)
        root.addStretch()

        row = QHBoxLayout(); row.addStretch()
        cancel = QPushButton("Cancel"); cancel.clicked.connect(self.reject); row.addWidget(cancel)
        self.action = QPushButton("Recover Account" if self.mode == "recover" else "Activate This Device")
        self.action.setObjectName("Primary")
        self.action.clicked.connect(self._submit)
        row.addWidget(self.action); root.addLayout(row)

    def _submit(self):
        email = self.email.text().strip()
        code = self.code.text().strip().upper()
        if not email or "@" not in email:
            QMessageBox.warning(self, "Purchase Email Required", "Enter the email address used for the original purchase.")
            return
        if not code:
            QMessageBox.warning(self, "Recovery Code Required", "Enter the Recovery Code issued for this license.")
            return

        self.action.setEnabled(False)
        self.action.setText("Verifying…")
        QApplication.processEvents()
        try:
            client = LicensingBackendClient(timeout_seconds=12.0)
            result = client.recover_existing_license(
                email=email,
                recovery_code=code,
                mode=self.mode,
            )
        except BackendClientError as exc:
            QMessageBox.warning(self, "Activation Failed", str(exc))
            return
        finally:
            self.action.setEnabled(True)
            self.action.setText("Recover Account" if self.mode == "recover" else "Activate This Device")

        # The server stores only the Recovery Code hash. Preserve the code supplied by
        # the customer locally so the License page can continue to display/copy it.
        result = dict(result)
        result["recovery_code"] = code
        parent = self.parent()
        if parent is not None and hasattr(parent, "apply_license_activation"):
            parent.apply_license_activation(result)

        migration_text = ""
        if result.get("migration_counted"):
            if str(result.get("license_type", "")).upper() == "PERSONAL":
                remaining = result.get("migrations_remaining_12_months")
                migration_text = f"\n\nThis recovery counted as a device migration/rebinding. Remaining in the current 12-month period: {remaining}."
            else:
                migration_text = "\n\nThis device rebind was recorded. Company licenses do not have an annual migration-count limit."
        else:
            migration_text = "\n\nSame-device restoration detected; no migration/rebinding was counted."

        QMessageBox.information(
            self,
            "License Activated",
            f"License verified and activated on this device.\n\nLicense ID: {result.get('license_id','')}\nPurchase Email: {result.get('email','')}" + migration_text,
        )
        self.accept()


class RefundRequestDialog(QDialog):
    """Submit a refund request for manual review. Submission does not revoke the license."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Request a Refund")
        self.resize(680, 660)
        self.setModal(True)
        self.setStyleSheet(f"""
            QDialog {{ background:#F6F8F7; }}
            QLabel#Title {{ color:{BRAND_GREEN_DARK}; font-size:22px; font-weight:800; }}
            QLabel#Sub {{ color:#5F6F67; }}
            QLabel#Required {{ color:#8C2F24; font-size:12px; font-weight:700; }}
            QLineEdit, QComboBox, QTextEdit {{ background:white; border:1px solid #CDD8D2; border-radius:7px; padding:8px; }}
            QPushButton#Primary {{ background:{BRAND_GREEN};color:white;border:none;border-radius:8px;padding:9px 14px;font-weight:700; }}
            QPushButton {{ min-height:34px; }}
        """)
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 24)
        root.setSpacing(8)

        title = QLabel("Request a Refund")
        title.setObjectName("Title")
        root.addWidget(title)
        sub = QLabel("Submit the purchase details for review. Your license remains active while the request is reviewed; revocation occurs only after a confirmed refund.")
        sub.setObjectName("Sub")
        sub.setWordWrap(True)
        root.addWidget(sub)

        required = QLabel("Fields marked * are required before a refund request can be submitted.")
        required.setObjectName("Required")
        required.setWordWrap(True)
        root.addWidget(required)

        root.addWidget(QLabel("Purchase Email *"))
        self.email = QLineEdit(); self.email.setPlaceholderText("Email address used for the purchase")
        root.addWidget(self.email)
        parent_state = getattr(parent, "license_state", {}) if parent is not None else {}
        self.email.setText(str(parent_state.get("email") or parent_state.get("purchase_email") or ""))

        root.addWidget(QLabel("Activation / Recovery Code *"))
        self.code = QLineEdit(); self.code.setPlaceholderText("Activation / Recovery Code")
        root.addWidget(self.code)
        self.code.setText(str(parent_state.get("recovery_code") or ""))
        self.no_code = QCheckBox("I no longer have access to my Activation / Recovery Code")
        self.no_code.toggled.connect(self._toggle_code_requirement)
        root.addWidget(self.no_code)

        root.addWidget(QLabel("Receipt / Purchase ID *"))
        self.receipt = QLineEdit(); self.receipt.setPlaceholderText("Receipt or Purchase ID from the confirmation email")
        root.addWidget(self.receipt)
        self.receipt.setText(str(parent_state.get("purchase_id") or parent_state.get("purchase_intent_id") or ""))

        root.addWidget(QLabel("Refund Reason *"))
        self.reason = QComboBox()
        self.reason.addItems([
            "Select a reason...",
            "Software does not operate as expected",
            "Technical problem",
            "Compatibility issue",
            "Duplicate purchase",
            "Incorrect license purchased",
            "Change of mind",
            "Other",
        ])
        root.addWidget(self.reason)

        root.addWidget(QLabel("Details *"))
        self.details = QTextEdit()
        self.details.setPlaceholderText("Briefly describe the reason for the refund request and any relevant details.")
        self.details.setMinimumHeight(100)
        root.addWidget(self.details)

        note = QLabel(
            "Australian Consumer Law rights are not excluded. Submitting a request does not immediately revoke the license. "
            "The request is reviewed first. If approved, the refund is processed through the original payment provider and the license is revoked only after the refund is confirmed."
        )
        note.setWordWrap(True)
        note.setObjectName("Sub")
        root.addWidget(note)

        row = QHBoxLayout(); row.addStretch()
        cancel = QPushButton("Cancel"); cancel.clicked.connect(self.reject); row.addWidget(cancel)
        submit = QPushButton("Submit Refund Request")
        submit.setObjectName("Primary")
        submit.clicked.connect(self._submit)
        row.addWidget(submit)
        root.addLayout(row)

    def _toggle_code_requirement(self, checked):
        self.code.setEnabled(not checked)
        if checked:
            self.code.clear()
            self.code.setPlaceholderText("Not available — purchase email and receipt ID will be required")
        else:
            self.code.setPlaceholderText("Activation / Recovery Code")

    def _submit(self):
        missing = []
        email = self.email.text().strip()
        receipt = self.receipt.text().strip()
        code = self.code.text().strip()
        details = self.details.toPlainText().strip()

        if not email:
            missing.append("Purchase Email")
        if not receipt:
            missing.append("Receipt / Purchase ID")
        if not self.no_code.isChecked() and not code:
            missing.append("Activation / Recovery Code")
        if self.reason.currentIndex() == 0:
            missing.append("Refund Reason")
        if not details:
            missing.append("Details")

        if missing:
            QMessageBox.warning(
                self,
                "Required Information Missing",
                "Please complete the following required information before submitting:\n\n• " + "\n• ".join(missing)
            )
            return

        self.setEnabled(False)
        QApplication.processEvents()
        try:
            client = LicensingBackendClient(timeout_seconds=12.0)
            result = client.submit_refund_request(
                email=email,
                recovery_code=code,
                no_recovery_code=self.no_code.isChecked(),
                purchase_reference=receipt,
                reason=self.reason.currentText(),
                details=details,
            )
        except BackendClientError as exc:
            QMessageBox.warning(self, "Refund Request Not Submitted", str(exc))
            return
        finally:
            self.setEnabled(True)

        request_id = str(result.get("refund_request_id", "")).strip()
        QMessageBox.information(
            self,
            "Refund Request Submitted",
            f"Refund request {request_id} has been recorded for review.\n\n"
            "Your license remains ACTIVE while the request is pending. It will be revoked only after the payment provider confirms an approved refund."
        )
        self.accept()

