from PyQt6.QtWidgets import QWidget,QVBoxLayout,QHBoxLayout,QLabel,QPushButton,QProgressBar,QFrame,QMessageBox
from ui.widgets.log_view import LogView

class BaseModulePage(QWidget):
    def __init__(self,info,task_manager,logger,back_callback):
        super().__init__(); self.info=info; self.task_manager=task_manager; self.logger=logger
        self.main=QVBoxLayout(self); self.main.setContentsMargins(32,24,32,24); self.main.setSpacing(10)
        top=QHBoxLayout(); back=QPushButton("← Report Center"); back.setObjectName("SecondaryButton"); back.clicked.connect(back_callback); top.addWidget(back); top.addSpacing(12)
        titles=QVBoxLayout(); title=QLabel(info.title.upper()); title.setObjectName("PageTitle"); subtitle=QLabel(info.subtitle); subtitle.setObjectName("PageSubtitle"); titles.addWidget(title); titles.addWidget(subtitle); top.addLayout(titles); top.addStretch(); self.main.addLayout(top)
        desc=QLabel(info.description); desc.setWordWrap(True); desc.setObjectName("PageSubtitle"); self.main.addWidget(desc)
        self.form_card=QFrame(); self.form_card.setObjectName("Panel"); self.form_layout=QVBoxLayout(self.form_card); self.form_layout.setContentsMargins(20,16,20,16); self.form_layout.setSpacing(8); self.main.addWidget(self.form_card)
        actions=QHBoxLayout(); actions.addStretch(); self.run_button=QPushButton("Run"); self.run_button.setObjectName("PrimaryButton"); self.run_button.setMinimumWidth(180); actions.addWidget(self.run_button); self.main.addLayout(actions)
        prog_title=QLabel("PROGRESS"); prog_title.setObjectName("SectionTitle"); self.main.addWidget(prog_title); self.progress=QProgressBar(); self.progress.setValue(0); self.main.addWidget(self.progress); self.status=QLabel("Ready"); self.status.setObjectName("Muted"); self.main.addWidget(self.status)
        log_title=QLabel("ACTIVITY LOG"); log_title.setObjectName("SectionTitle"); self.main.addWidget(log_title); self.log_view=LogView(); self.main.addWidget(self.log_view,1)
    def start_task(self,callable_):
        self.run_button.setEnabled(False); self.progress.setValue(0); self.status.setText("Running…"); self.log_view.add("Task started")
        self.task_manager.start(callable_,{"progress":self._on_progress,"log":self._on_log,"completed":self._on_completed,"failed":self._on_failed})
    def _on_progress(self,value,message): self.progress.setValue(value); self.status.setText(message or f"{value}%")
    def _on_log(self,message): self.log_view.add(message); self.logger.info(message)
    def _on_completed(self,result):
        self.run_button.setEnabled(True); self.progress.setValue(100); self.status.setText(result.message); self.log_view.add(result.message)
        for p in getattr(result,"output_files",[]): self.log_view.add(f"Output: {p}")
        QMessageBox.information(self,"Task Completed",result.message)
    def _on_failed(self,message):
        self.run_button.setEnabled(True); self.status.setText("Failed"); self.log_view.add(f"ERROR: {message}"); self.logger.error(message); QMessageBox.critical(self,"Task Failed",message)
