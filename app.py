import sys
from pathlib import Path
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import QTimer
from PyQt6.QtGui import QIcon
from ui.styles import PUBLIC_STYLE

from core.application.module_manager import ModuleManager
from core.application.task_manager import TaskManager
from core.services.config_service import ConfigService
from core.services.logging_service import LoggingService
from core.services.trial_service import TrialService
from core.utils.app_paths import logs_dir
from modules.split_reports.module import SplitReportsModule
from modules.consolidate_reports.module import ConsolidateReportsModule
from modules.synchronize_reports.module import SynchronizeReportsModule
from modules.lookup_data.module import LookupDataModule
from ui.main_window import MainWindow
from ui.startup_splash import PublishStartupSplash
from version import APP_NAME


class MasterReportCenterApplication:
    def __init__(self):
        self.qt_app = QApplication(sys.argv)
        self.qt_app.setApplicationName(APP_NAME)
        self.base_dir = Path(__file__).resolve().parent
        self.config = ConfigService(self.base_dir / "config")
        self.logger = LoggingService(logs_dir())
        self.task_manager = TaskManager(self.logger)
        self.trial_service = TrialService()
        self.module_manager = ModuleManager()
        self._register_modules()
        self._apply_theme()

    def _register_modules(self):
        self.module_manager.register(SplitReportsModule())
        self.module_manager.register(ConsolidateReportsModule())
        self.module_manager.register(SynchronizeReportsModule())
        self.module_manager.register(LookupDataModule())

    def _apply_theme(self):
        self.qt_app.setStyleSheet(PUBLIC_STYLE)
        icon = self.base_dir / "assets" / "app.ico"
        if icon.exists():
            self.qt_app.setWindowIcon(QIcon(str(icon)))

    def run(self):
        splash = PublishStartupSplash(self.base_dir / "assets")
        screen = self.qt_app.primaryScreen()
        if screen is not None:
            splash.show_centered(screen)
        else:
            splash.show()
        self.qt_app.processEvents()

        window = MainWindow(
            module_manager=self.module_manager,
            task_manager=self.task_manager,
            config_service=self.config,
            logger=self.logger,
            trial_service=self.trial_service,
        )

        def finish_startup():
            splash.close()
            window.show()
            window.raise_()
            window.activateWindow()

        QTimer.singleShot(1250, finish_startup)
        return self.qt_app.exec()
