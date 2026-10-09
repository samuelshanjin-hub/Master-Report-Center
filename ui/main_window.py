import json, os, shutil
from pathlib import Path
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QPixmap
from PyQt6.QtWidgets import QMainWindow,QWidget,QVBoxLayout,QHBoxLayout,QLabel,QPushButton,QFrame,QStackedWidget,QStatusBar,QFileDialog,QMessageBox
from ui.pages.home_page import HomePage, MODULE_PRESENTATION
from ui.pages.simple_pages import SettingsPage, LicensePage, AboutPage
from version import VERSION
from ui.payment_dialog import PaymentDialog, ActivationDialog, RefundRequestDialog
from core.services.backend_client import LicensingBackendClient, BackendClientError
from core.utils.app_paths import logs_dir, user_config_dir, default_documents_dir

class MainWindow(QMainWindow):
    def __init__(self,module_manager,task_manager,config_service,logger,trial_service=None):
        super().__init__(); self.module_manager=module_manager; self.task_manager=task_manager; self.config_service=config_service; self.logger=logger; self.trial_service=trial_service
        self.base_dir=Path(__file__).resolve().parents[1]; self.assets=self.base_dir/"assets"; self.nav_buttons={}; self.pages={}
        self.license_state = trial_service.get_state() if trial_service is not None else {"status":"trial","days_remaining":12,"license_type":None}
        self.setWindowTitle("Golden Fir Report Center"); self.resize(1500,900); self.setMinimumSize(1024,640)
        central=QWidget(); self.setCentralWidget(central); root=QVBoxLayout(central); root.setContentsMargins(0,0,0,0); root.setSpacing(0)
        root.addWidget(self._build_topbar()); body=QHBoxLayout(); body.setContentsMargins(0,0,0,0); body.setSpacing(0); root.addLayout(body,1)
        body.addWidget(self._build_sidebar()); self.stack=QStackedWidget(); self.stack.setObjectName("ContentArea"); body.addWidget(self.stack,1)
        self._register_static_pages(); self.setStatusBar(QStatusBar()); self.set_status("Ready"); self.navigate("home"); QTimer.singleShot(900, lambda: self.refresh_remote_license_status(silent=True))

    def _pix(self,name,w,h):
        p=self.assets/name
        if not p.exists(): return QPixmap()
        return QPixmap(str(p)).scaled(w,h,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation)

    def _build_topbar(self):
        top=QFrame(); top.setObjectName("TopBar"); tl=QHBoxLayout(top); tl.setContentsMargins(16,8,16,8)
        mark=QLabel(); mark.setPixmap(self._pix("brand_mark.png",46,32)); tl.addWidget(mark)
        title=QLabel("Golden Fir Report Center"); title.setObjectName("HeaderBrand"); tl.addWidget(title); tl.addStretch()
        version=QLabel(f"Version {VERSION}"); version.setObjectName("Muted"); tl.addWidget(version); return top

    def _build_sidebar(self):
        side=QFrame(); side.setObjectName("Sidebar"); side.setMinimumWidth(220); side.setMaximumWidth(260); sl=QVBoxLayout(side); sl.setContentsMargins(14,20,14,16); sl.setSpacing(7)
        self.brand_logo=QLabel(); self.brand_logo.setPixmap(self._pix("golden_fir_wordmark.png",225,125)); self.brand_logo.setAlignment(Qt.AlignmentFlag.AlignCenter); sl.addWidget(self.brand_logo)
        self.module_logo=QLabel(); self.module_logo.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop); self.module_logo.setMinimumHeight(120); self.module_logo.setMaximumHeight(120); self.module_logo.hide(); sl.addWidget(self.module_logo)
        self.module_title=QLabel(); self.module_title.setAlignment(Qt.AlignmentFlag.AlignCenter); self.module_title.hide(); sl.addWidget(self.module_title)
        self.module_description=QLabel(); self.module_description.setWordWrap(True); self.module_description.setAlignment(Qt.AlignmentFlag.AlignCenter); self.module_description.hide(); sl.addWidget(self.module_description)
        line=QFrame(); line.setFrameShape(QFrame.Shape.HLine); line.setObjectName("Divider"); sl.addWidget(line)
        navs=[("⌂  Home","home"),("♢  License","license"),("⚙  Settings","settings"),("ⓘ  About","about")]
        for text,key in navs:
            b=QPushButton(text); b.setObjectName("Nav"); b.clicked.connect(lambda checked=False,k=key:self.navigate(k)); sl.addWidget(b); self.nav_buttons[key]=b
        sl.addStretch(); company=QLabel("Golden Fir Solutions"); company.setAlignment(Qt.AlignmentFlag.AlignCenter); company.setStyleSheet("color:#B58A2A;font-size:14px;font-weight:800;"); sl.addWidget(company); foot=QLabel("Empowering your reporting workflow"); foot.setAlignment(Qt.AlignmentFlag.AlignCenter); foot.setObjectName("Muted"); sl.addWidget(foot); return side

    def _register_static_pages(self):
        self.pages={"home":HomePage(self,self.module_manager.all()),"license":LicensePage(self),"settings":SettingsPage(self),"about":AboutPage()}
        for p in self.pages.values(): self.stack.addWidget(p)

    def _set_nav_active(self,key):
        for k,b in self.nav_buttons.items(): b.setObjectName("NavActive" if k==key else "Nav"); b.style().unpolish(b); b.style().polish(b)

    def clear_module_info(self):
        self.module_logo.hide(); self.module_title.hide(); self.module_description.hide(); self.brand_logo.show()

    def set_module_info(self,module_id):
        p=MODULE_PRESENTATION.get(module_id); self.brand_logo.hide()
        if p:
            self.module_logo.setPixmap(self._pix(p[3],145,115)); self.module_logo.show()
            self.module_title.setText(p[1]); self.module_title.setStyleSheet(
                f"color:{p[5]};font-size:15px;font-weight:700;"
            ); self.module_title.show()
            self.module_description.setText(p[2].replace("\n"," ")); self.module_description.setStyleSheet(
                "color:#66776E;font-size:12px;"
            ); self.module_description.show()

    def navigate(self,key):
        if key not in self.pages: return
        if key == "license":
            self.refresh_remote_license_status(silent=True)
        self.stack.setCurrentWidget(self.pages[key]); self.clear_module_info(); self._set_nav_active(key); self.set_status(key.replace("_"," ").title())

    def open_module(self,module_id):
        current_status = str(self.license_state.get("status", "trial")).lower()
        if current_status in {"refunded", "revoked"}:
            QMessageBox.warning(
                self,
                "License Unavailable",
                "This license is no longer active. The report modules are locked. Existing reports and files have not been deleted."
            )
            self.navigate("license")
            return
        if current_status == "expired":
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Information)
            box.setWindowTitle("Trial Expired")
            box.setText("Your 12-day trial has ended.")
            box.setInformativeText("Choose Buy License to purchase, or Activate / Recover if you already have a license.")
            buy_button = box.addButton("Buy License", QMessageBox.ButtonRole.AcceptRole)
            account_button = box.addButton("Activate / Recover", QMessageBox.ButtonRole.ActionRole)
            box.addButton("Close", QMessageBox.ButtonRole.RejectRole)
            box.exec()
            clicked = box.clickedButton()
            if clicked is buy_button:
                self.open_payment_preview()
            elif clicked is account_button:
                self.navigate("license")
            return
        if module_id not in self.pages:
            module=self.module_manager.get(module_id); page=module.create_page(self.task_manager,self.logger,lambda:self.navigate("home")); self.pages[module_id]=page; self.stack.addWidget(page)
        self.stack.setCurrentWidget(self.pages[module_id]); self.set_module_info(module_id); self._set_nav_active(""); self.set_status(self.module_manager.get(module_id).info.title)


    def open_payment_preview(self):
        PaymentDialog(self.assets, self).exec()

    def open_activation_preview(self):
        ActivationDialog("Activate License", "activate", self).exec()

    def open_recovery_preview(self):
        ActivationDialog("Recover Existing Account", "recover", self).exec()

    def open_refund_preview(self):
        RefundRequestDialog(self).exec()

    def apply_license_activation(self, payload):
        raw_type = str(payload.get("license_type", "PERSONAL")).upper()
        license_type = "Company License" if raw_type == "COMPANY" else "Personal License"
        state = {
            "status": "active",
            "days_remaining": 0,
            "license_type": license_type,
            "email": str(payload.get("email", "")).strip(),
            "purchase_email": str(payload.get("email", "")).strip(),
            "purchase_id": str(payload.get("purchase_id", "")).strip(),
            "purchase_intent_id": str(payload.get("purchase_intent_id", "")).strip(),
            "license_id": str(payload.get("license_id", "")).strip(),
            "recovery_code": str(payload.get("recovery_code", "") or "").strip(),
            "recovery_code_last4": str(payload.get("recovery_code_last4", "") or "").strip(),
            "activated_at": str(payload.get("activated_at", "") or "").strip(),
            "same_device": payload.get("same_device"),
            "migration_counted": payload.get("migration_counted"),
            "migrations_used_12_months": payload.get("migrations_used_12_months"),
            "migrations_remaining_12_months": payload.get("migrations_remaining_12_months"),
        }
        # A newly issued Personal license has not consumed any migration/rebinding allowance yet.
        if raw_type == "PERSONAL" and state.get("migrations_used_12_months") is None:
            state["migrations_used_12_months"] = 0
            state["migrations_remaining_12_months"] = 4
        self.license_state = state
        if self.trial_service is not None and hasattr(self.trial_service, "save_active_license"):
            self.trial_service.save_active_license(state)
        home = self.pages.get("home")
        if hasattr(home, "set_license_state"):
            home.set_license_state(self.license_state)
        license_page = self.pages.get("license")
        if hasattr(license_page, "set_license_state"):
            license_page.set_license_state(self.license_state)
        self.set_status(f"License active — {license_type}")

    def refresh_remote_license_status(self, silent=True):
        state = dict(self.license_state or {})
        status = str(state.get("status", "")).lower()
        license_id = str(state.get("license_id", "") or "").strip()
        if status not in {"active", "refunded", "revoked"} or not license_id:
            return
        try:
            remote = LicensingBackendClient(timeout_seconds=5.0).get_license_status(license_id)
        except BackendClientError as exc:
            if not silent:
                QMessageBox.warning(self, "License Status", str(exc))
            return

        remote_status = str(remote.get("status", "")).lower()
        state["migrations_used_12_months"] = remote.get("migrations_used_12_months")
        state["migrations_remaining_12_months"] = remote.get("migrations_remaining_12_months")
        state["refund_request_status"] = remote.get("refund_request_status")
        if remote_status in {"refunded", "revoked"}:
            state["status"] = remote_status
            state["recovery_code"] = ""
        else:
            state["status"] = "active"
        self.license_state = state
        if self.trial_service is not None and hasattr(self.trial_service, "save_license_state"):
            self.trial_service.save_license_state(state)
        home = self.pages.get("home")
        if hasattr(home, "set_license_state"):
            home.set_license_state(state)
        license_page = self.pages.get("license")
        if hasattr(license_page, "set_license_state"):
            license_page.set_license_state(state)
        if remote_status in {"refunded", "revoked"}:
            self.set_status(f"License {remote_status}")

    def set_license_active(self, license_type="Personal License"):
        """Compatibility helper for old callers; does not create a paid entitlement."""
        self.set_status("Activation requires backend verification.")

    def manage_license(self):
        self.navigate("license")

    def set_status(self,text): self.statusBar().showMessage(text)
    def _open_folder(self,path):
        try:
            if os.name=="nt": os.startfile(str(path))
            else: QMessageBox.information(self,"Folder",str(path))
        except Exception as exc: QMessageBox.warning(self,"Open Folder",str(exc))
    def open_log_folder(self):
        p=logs_dir(); p.mkdir(parents=True,exist_ok=True); self._open_folder(p)
    def import_configuration(self):
        src,_=QFileDialog.getOpenFileName(self,"Import Configuration",str(default_documents_dir()),"JSON Files (*.json)")
        if not src:return
        try:
            data=json.loads(Path(src).read_text(encoding="utf-8")); cfg_dir=user_config_dir(); cfg_dir.mkdir(parents=True,exist_ok=True); dst=cfg_dir/"imported_configuration.json"; dst.write_text(json.dumps(data,indent=2,ensure_ascii=False),encoding="utf-8"); self.set_status(f"Configuration imported: {Path(src).name}"); QMessageBox.information(self,"Import Configuration",f"Imported to:\n{dst}")
        except Exception as exc: QMessageBox.critical(self,"Import Configuration",str(exc))
    def export_configuration(self):
        dst,_=QFileDialog.getSaveFileName(self,"Export Configuration",str(default_documents_dir()/"Master_Report_Center_Config.json"),"JSON Files (*.json)")
        if not dst:return
        try:
            payload={"app":self.config_service.app_config,"defaults":self.config_service.defaults,"version":VERSION}; Path(dst).write_text(json.dumps(payload,indent=2,ensure_ascii=False),encoding="utf-8"); self.set_status("Configuration exported"); QMessageBox.information(self,"Export Configuration",f"Saved:\n{dst}")
        except Exception as exc: QMessageBox.critical(self,"Export Configuration",str(exc))
    def check_for_updates(self): QMessageBox.information(self,"Check for Updates","Automatic update checking is not enabled in this version. Please use the official Golden Fir release page or Microsoft Store when available.")
