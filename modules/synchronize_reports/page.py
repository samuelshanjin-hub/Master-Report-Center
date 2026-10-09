from __future__ import annotations

import os
from pathlib import Path

from PyQt6.QtCore import Qt, QRect
from PyQt6.QtGui import QColor, QPainter
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QInputDialog,
    QComboBox,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QTextEdit,
    QVBoxLayout,
    QBoxLayout,
    QWidget,
)

from .service import (
    HeaderAnalysis,
    SynchronizeReportsService,
)


class DualColorProgressBar(QProgressBar):
    """Same dual-colour percentage treatment used by Split Reports."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setTextVisible(False)

    def paintEvent(self, event):
        super().paintEvent(event)

        minimum = self.minimum()
        maximum = self.maximum()
        value = self.value()

        if maximum <= minimum:
            ratio = 0.0
        else:
            ratio = (
                (value - minimum)
                / (maximum - minimum)
            )
            ratio = max(
                0.0,
                min(1.0, ratio),
            )

        text = f"{value}%"
        rect = self.rect()

        painter = QPainter(self)
        painter.setRenderHint(
            QPainter.RenderHint.TextAntialiasing,
            True,
        )

        painter.setPen(
            QColor("#17231D")
        )
        painter.drawText(
            rect,
            int(Qt.AlignmentFlag.AlignCenter),
            text,
        )

        filled_width = int(
            rect.width() * ratio
        )

        if filled_width > 0:
            painter.save()
            painter.setClipRect(
                QRect(
                    0,
                    0,
                    filled_width,
                    rect.height(),
                )
            )
            painter.setPen(
                QColor("#FFFFFF")
            )
            painter.drawText(
                rect,
                int(Qt.AlignmentFlag.AlignCenter),
                text,
            )
            painter.restore()

        painter.end()


class SynchronizeReportsPage(QWidget):
    """
    Publish-style Synchronize Reports page connected to Core V2 TaskManager.
    """

    def __init__(
        self,
        info,
        task_manager,
        logger,
        back_callback,
    ):
        super().__init__()

        self.info = info
        self.task_manager = task_manager
        self.logger = logger
        self.back_callback = back_callback

        self.service = SynchronizeReportsService()
        self.analysis: HeaderAnalysis | None = None
        self._worksheet_password_cache = None
        self._destination_sheet_items = []
        self._source_sheet_items = []

        self._build()

    def _build(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.page_scroll = QScrollArea()
        self.page_scroll.setObjectName("ModuleScroll")
        self.page_scroll.setWidgetResizable(True)
        self.page_scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        self.page_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.page_scroll.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )
        outer.addWidget(self.page_scroll)

        self.page_content = QWidget()
        self.page_content.setObjectName("ModuleScrollContent")
        self.page_scroll.setWidget(self.page_content)

        root = QVBoxLayout(self.page_content)
        root.setContentsMargins(
            22,
            18,
            22,
            18,
        )
        root.setSpacing(10)

        top = QHBoxLayout()

        back = QPushButton(
            "← Report Center"
        )
        back.setObjectName(
            "Secondary"
        )
        back.clicked.connect(
            self.back_callback
        )

        title = QLabel(
            "Synchronize Reports"
        )
        title.setObjectName(
            "PageTitle"
        )

        top.addWidget(back)
        top.addWidget(title)
        top.addStretch(1)

        root.addLayout(top)

        subtitle = QLabel(
            "Match source rows to a destination workbook and update selected fields."
        )
        subtitle.setObjectName(
            "Muted"
        )
        root.addWidget(subtitle)

        self.progress = DualColorProgressBar()
        self.progress.setRange(
            0,
            100,
        )
        self.progress.setValue(0)
        root.addWidget(
            self.progress
        )

        self.work = QHBoxLayout()
        self.work.setSpacing(12)

        self.config = QWidget()
        self.grid = QGridLayout(self.config)
        self.grid.setContentsMargins(
            0,
            0,
            0,
            0,
        )
        self.grid.setSpacing(10)

        # 1. Destination
        self.destination_group = QGroupBox(
            "1. Destination"
        )
        destination_layout = QGridLayout(
            self.destination_group
        )

        self.destination_edit = QLineEdit()
        self.destination_edit.setPlaceholderText(
            "Select the Destination / Master workbook"
        )

        destination_browse = QPushButton(
            "Browse..."
        )
        destination_browse.setObjectName(
            "Secondary"
        )
        destination_browse.clicked.connect(
            self.select_destination
        )

        destination_open = QPushButton(
            "Open Excel"
        )
        destination_open.setObjectName(
            "Green"
        )
        destination_open.clicked.connect(
            lambda: self.open_excel(
                self.destination_edit.text().strip()
            )
        )

        self.destination_sheet = QComboBox()
        self.destination_sheet.setToolTip(
            "Select the Destination worksheet to update. Hidden and VeryHidden "
            "worksheets are shown with their visibility state."
        )
        self.destination_sheet.currentIndexChanged.connect(
            self._destination_sheet_changed
        )

        destination_layout.addWidget(
            QLabel("Destination Excel File"),
            0,
            0,
            1,
            3,
        )
        destination_layout.addWidget(
            self.destination_edit,
            1,
            0,
        )
        destination_layout.addWidget(
            destination_browse,
            1,
            1,
        )
        destination_layout.addWidget(
            destination_open,
            1,
            2,
        )
        destination_layout.addWidget(
            QLabel("Destination Worksheet"),
            2,
            0,
        )
        destination_layout.addWidget(
            self.destination_sheet,
            2,
            1,
            1,
            2,
        )

        # 2. Source
        self.source_group = QGroupBox(
            "2. Source"
        )
        source_layout = QGridLayout(
            self.source_group
        )

        self.source_edit = QLineEdit()
        self.source_edit.setPlaceholderText(
            "Select folder containing Source reports"
        )

        source_browse = QPushButton(
            "Browse..."
        )
        source_browse.setObjectName(
            "Secondary"
        )
        source_browse.clicked.connect(
            self.select_source_folder
        )

        source_open = QPushButton(
            "Open Folder"
        )
        source_open.setObjectName(
            "Info"
        )
        source_open.clicked.connect(
            lambda: self.open_folder(
                self.source_edit.text().strip()
            )
        )

        self.source_sheet = QComboBox()
        self.source_sheet.setToolTip(
            "Worksheets are read from the first valid Source workbook as a sample. "
            "The selected worksheet name must exist in each Source workbook."
        )
        self.source_sheet.currentIndexChanged.connect(
            self._source_sheet_changed
        )

        source_layout.addWidget(
            QLabel("Source Reports Folder"),
            0,
            0,
            1,
            3,
        )
        source_layout.addWidget(
            self.source_edit,
            1,
            0,
        )
        source_layout.addWidget(
            source_browse,
            1,
            1,
        )
        source_layout.addWidget(
            source_open,
            1,
            2,
        )
        source_layout.addWidget(
            QLabel("Source Worksheet"),
            2,
            0,
        )
        source_layout.addWidget(
            self.source_sheet,
            2,
            1,
            1,
            2,
        )

        # 3. Field Selection
        self.fields_group = QGroupBox(
            "3. Field Selection"
        )
        fields_layout = QGridLayout(
            self.fields_group
        )

        field_note = QLabel(
            "Analyze Headers first. Match fields should be stable identifiers; "
            "Write fields are values that may change."
        )
        field_note.setWordWrap(True)
        field_note.setObjectName(
            "Muted"
        )

        self.match_list = QListWidget()
        self.match_list.setSelectionMode(
            QAbstractItemView.SelectionMode.ExtendedSelection
        )

        self.write_list = QListWidget()
        self.write_list.setSelectionMode(
            QAbstractItemView.SelectionMode.ExtendedSelection
        )

        fields_layout.addWidget(
            field_note,
            0,
            0,
            1,
            2,
        )
        fields_layout.addWidget(
            QLabel("Match Fields"),
            1,
            0,
        )
        fields_layout.addWidget(
            QLabel("Write Fields"),
            1,
            1,
        )
        fields_layout.addWidget(
            self.match_list,
            2,
            0,
        )
        fields_layout.addWidget(
            self.write_list,
            2,
            1,
        )

        # 4. Options
        self.options_group = QGroupBox(
            "4. Options"
        )
        options_layout = QVBoxLayout(
            self.options_group
        )

        self.append_unmatched = QCheckBox(
            "Add unmatched source rows"
        )
        self.append_unmatched.setChecked(
            True
        )

        self.ignore_blank_updates = QCheckBox(
            "Ignore blank source values when updating matched rows"
        )
        self.ignore_blank_updates.setChecked(
            True
        )

        options_layout.addWidget(
            self.append_unmatched
        )
        options_layout.addWidget(
            self.ignore_blank_updates
        )

        duplicate_note = QLabel(
            "If duplicate Match keys exist in the Destination, the first "
            "Destination row is used and the duplicate-key count is reported."
        )
        duplicate_note.setWordWrap(True)
        duplicate_note.setObjectName(
            "Muted"
        )
        options_layout.addWidget(
            duplicate_note
        )

        # Start
        self.actions_group = QGroupBox(
            "Start Synchronization"
        )
        self.actions_layout = QBoxLayout(
            QBoxLayout.Direction.LeftToRight,
            self.actions_group,
        )
        actions_layout = self.actions_layout

        self.analyze_button = QPushButton(
            "Analyze Headers"
        )
        self.analyze_button.setObjectName(
            "Info"
        )
        self.analyze_button.clicked.connect(
            self.start_analysis
        )

        self.start_button = QPushButton(
            "▶ Start Synchronization"
        )
        self.start_button.setObjectName(
            "Primary"
        )
        self.start_button.setEnabled(
            False
        )
        self.start_button.clicked.connect(
            self.start_synchronization
        )

        actions_layout.addWidget(
            self.analyze_button
        )
        actions_layout.addWidget(
            self.start_button
        )

        self.grid.addWidget(
            self.destination_group,
            0,
            0,
        )
        self.grid.addWidget(
            self.source_group,
            0,
            1,
        )
        self.grid.addWidget(
            self.fields_group,
            1,
            0,
            1,
            2,
        )
        self.grid.addWidget(
            self.options_group,
            2,
            0,
        )
        self.grid.addWidget(
            self.actions_group,
            2,
            1,
        )

        self.log_group = QGroupBox(
            "Activity Log"
        )
        log_layout = QVBoxLayout(
            self.log_group
        )

        clear = QPushButton(
            "Clear"
        )
        clear.setObjectName(
            "Secondary"
        )
        clear.clicked.connect(
            self.log_clear
        )
        log_layout.addWidget(
            clear,
            0,
            Qt.AlignmentFlag.AlignRight,
        )

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(320)
        self.log.setPlaceholderText(
            "Header analysis and synchronization activity will appear here."
        )

        log_layout.addWidget(
            self.log
        )

        self.work.addWidget(
            self.config,
            3,
        )
        self.work.addWidget(
            self.log_group,
            2,
        )

        root.addLayout(
            self.work,
            1,
        )

        self._compact_layout = None
        self._apply_responsive_layout(force=True)

    def _clear_grid(self):
        while self.grid.count():
            self.grid.takeAt(0)

    def _apply_responsive_layout(self, force: bool = False):
        if not hasattr(self, "page_scroll"):
            return

        width = self.page_scroll.viewport().width()
        compact = width < 1380

        if not force and compact == self._compact_layout:
            return

        self._clear_grid()

        if compact:
            self.grid.addWidget(self.destination_group, 0, 0)
            self.grid.addWidget(self.source_group, 1, 0)
            self.grid.addWidget(self.fields_group, 2, 0)
            self.grid.addWidget(self.options_group, 3, 0)
            self.grid.addWidget(self.actions_group, 4, 0)

            self.grid.setColumnStretch(0, 1)
            self.grid.setColumnStretch(1, 0)

            self.work.setDirection(QBoxLayout.Direction.TopToBottom)
            self.work.setStretch(0, 0)
            self.work.setStretch(1, 0)

            self.actions_layout.setDirection(
                QBoxLayout.Direction.TopToBottom
            )
            self.log.setMinimumHeight(360)
        else:
            self.grid.addWidget(self.destination_group, 0, 0)
            self.grid.addWidget(self.source_group, 0, 1)
            self.grid.addWidget(self.fields_group, 1, 0, 1, 2)
            self.grid.addWidget(self.options_group, 2, 0)
            self.grid.addWidget(self.actions_group, 2, 1)

            self.grid.setColumnStretch(0, 1)
            self.grid.setColumnStretch(1, 1)

            self.work.setDirection(QBoxLayout.Direction.LeftToRight)
            self.work.setStretch(0, 3)
            self.work.setStretch(1, 2)

            self.actions_layout.setDirection(
                QBoxLayout.Direction.LeftToRight
            )
            self.log.setMinimumHeight(320)

        self._compact_layout = compact

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._apply_responsive_layout()

    def _append_log(
        self,
        message: str,
    ):
        self.log.append(
            str(message)
        )

        try:
            self.logger.info(
                str(message)
            )
        except Exception:
            pass

    def log_clear(self):
        self.log.clear()

    def _invalidate_analysis(self):
        self.analysis = None
        self.match_list.clear()
        self.write_list.clear()
        self.start_button.setEnabled(
            False
        )

    @staticmethod
    def _worksheet_display_text(item: dict) -> str:
        name = str(item.get("name", ""))
        state = str(item.get("visibility", "Visible"))
        if state == "Visible":
            return name
        return f"{name} — {state}"

    @staticmethod
    def _select_default_visible(combo: QComboBox, items: list[dict]):
        if not items:
            return
        target_index = 0
        for index, item in enumerate(items):
            if item.get("visibility") == "Visible":
                target_index = index
                break
        combo.setCurrentIndex(target_index)

    def _destination_sheet_changed(self):
        self._worksheet_password_cache = None
        self._invalidate_analysis()

    def _source_sheet_changed(self):
        self._invalidate_analysis()

    def _selected_destination_sheet_name(self) -> str:
        return str(self.destination_sheet.currentData() or "").strip()

    def _selected_source_sheet_name(self) -> str:
        return str(self.source_sheet.currentData() or "").strip()

    def _load_destination_worksheets(self):
        self.destination_sheet.blockSignals(True)
        self.destination_sheet.clear()
        self._destination_sheet_items = []

        try:
            path = self.destination_edit.text().strip()
            if not path:
                return

            items = self.service.get_destination_worksheets(path)
            self._destination_sheet_items = list(items)

            for item in items:
                self.destination_sheet.addItem(
                    self._worksheet_display_text(item),
                    item["name"],
                )

            self._select_default_visible(
                self.destination_sheet,
                self._destination_sheet_items,
            )

        except Exception as exc:
            self._append_log(
                f"Unable to load Destination worksheets: {exc}"
            )

        finally:
            self.destination_sheet.blockSignals(False)

    def _load_source_worksheets(self):
        self.source_sheet.blockSignals(True)
        self.source_sheet.clear()
        self._source_sheet_items = []

        try:
            folder = self.source_edit.text().strip()
            if not folder:
                return

            destination = self.destination_edit.text().strip()
            result = self.service.get_source_sample_worksheets(
                source_folder=folder,
                destination_file=destination,
            )

            items = result["worksheets"]
            self._source_sheet_items = list(items)

            for item in items:
                self.source_sheet.addItem(
                    self._worksheet_display_text(item),
                    item["name"],
                )

            self._select_default_visible(
                self.source_sheet,
                self._source_sheet_items,
            )

            self._append_log(
                f"Source worksheet sample: {result['sample_file']}"
            )

        except Exception as exc:
            self._append_log(
                f"Unable to load Source worksheets: {exc}"
            )

        finally:
            self.source_sheet.blockSignals(False)

    def _get_destination_protection_password(self):
        destination_file = self.destination_edit.text().strip()
        destination_sheet = self._selected_destination_sheet_name()

        protection = self.service.get_destination_protection_info(
            destination_file=destination_file,
            destination_sheet=destination_sheet,
        )

        if not protection.get("is_protected", False):
            return protection, "", True

        if self._worksheet_password_cache is not None:
            return protection, self._worksheet_password_cache, True

        if self.service.validate_destination_worksheet_password(
            destination_file=destination_file,
            destination_sheet=protection["sheet_name"],
            password="",
        ):
            self._worksheet_password_cache = ""
            return protection, "", True

        password, accepted = QInputDialog.getText(
            self,
            "Worksheet Protection",
            (
                "The Destination worksheet is protected.\n\n"
                "A worksheet password is required to synchronize the report.\n"
                "The password is used only during this program session "
                "and is never saved or logged."
            ),
            QLineEdit.EchoMode.Password,
        )

        if not accepted:
            return protection, "", False

        if not self.service.validate_destination_worksheet_password(
            destination_file=destination_file,
            destination_sheet=protection["sheet_name"],
            password=password,
        ):
            QMessageBox.warning(
                self,
                "Incorrect Password",
                (
                    "The worksheet could not be unlocked with the password entered.\n\n"
                    "No source or Destination workbook changes were made."
                ),
            )
            return protection, "", False

        self._worksheet_password_cache = password
        return protection, password, True

    def select_destination(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Destination Workbook",
            self.destination_edit.text().strip()
            or str(Path.home()),
            "Excel Files (*.xlsx *.xlsm *.xlsb *.xls);;All Files (*.*)",
        )

        if path:
            self.destination_edit.setText(
                path
            )
            self._worksheet_password_cache = None
            self._invalidate_analysis()
            self._load_destination_worksheets()
            self._append_log(
                f"Destination selected: {path}"
            )

    def select_source_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Source Reports Folder",
            self.source_edit.text().strip()
            or str(Path.home()),
        )

        if folder:
            self.source_edit.setText(
                folder
            )
            self._invalidate_analysis()
            self._load_source_worksheets()
            self._append_log(
                f"Source folder selected: {folder}"
            )

    def open_excel(
        self,
        file_path: str,
    ):
        file_path = file_path.strip()

        if not file_path:
            return

        path = Path(file_path)

        if not path.is_file():
            QMessageBox.warning(
                self,
                "File not found",
                str(path),
            )
            return

        try:
            os.startfile(
                str(path)
            )
        except Exception as exc:
            QMessageBox.warning(
                self,
                "Unable to open Excel",
                str(exc),
            )

    def open_folder(
        self,
        folder: str,
    ):
        folder = folder.strip()

        if not folder:
            return

        path = Path(folder)

        if not path.is_dir():
            QMessageBox.warning(
                self,
                "Folder not found",
                str(path),
            )
            return

        try:
            os.startfile(
                str(path)
            )
        except Exception as exc:
            QMessageBox.warning(
                self,
                "Unable to open folder",
                str(exc),
            )

    def _paths_valid(self) -> bool:
        if not self.destination_edit.text().strip():
            QMessageBox.warning(
                self,
                "Missing Destination",
                "Please select the Destination workbook.",
            )
            return False

        if not self.source_edit.text().strip():
            QMessageBox.warning(
                self,
                "Missing Source Folder",
                "Please select the Source reports folder.",
            )
            return False

        return True

    def start_analysis(self):
        if not self._paths_valid():
            return

        self._invalidate_analysis()
        self.progress.setValue(0)
        self.log.clear()

        self.analyze_button.setEnabled(
            False
        )

        self._append_log(
            "Analyzing Destination and Source headers..."
        )

        if self.destination_sheet.count() == 0:
            self._load_destination_worksheets()

        if self.source_sheet.count() == 0:
            self._load_source_worksheets()

        destination_sheet = self._selected_destination_sheet_name()
        source_sheet = self._selected_source_sheet_name()

        if not destination_sheet or not source_sheet:
            QMessageBox.warning(
                self,
                "Worksheet Selection Required",
                "Select both a Destination worksheet and a Source worksheet.",
            )
            self.analyze_button.setEnabled(True)
            return

        self.task_manager.start(
            lambda context: self.service.analyze(
                destination_file=self.destination_edit.text().strip(),
                source_folder=self.source_edit.text().strip(),
                destination_sheet=destination_sheet,
                source_sheet=source_sheet,
                context=context,
            ),
            {
                "progress": self.on_progress,
                "log": self._append_log,
                "completed": self.on_analysis_complete,
                "failed": self.on_analysis_failed,
            },
        )

    def on_analysis_complete(
        self,
        result: HeaderAnalysis,
    ):
        self.analysis = result

        self.match_list.clear()
        self.write_list.clear()

        for header in result.common_headers:
            self.match_list.addItem(
                header
            )
            self.write_list.addItem(
                header
            )

        self.progress.setValue(
            100
        )

        self.analyze_button.setEnabled(
            True
        )
        self.start_button.setEnabled(
            True
        )

        self._append_log("")
        self._append_log(
            "========== Header Analysis =========="
        )
        self._append_log(
            f"Destination sheet: {result.destination_sheet}"
        )
        self._append_log(
            f"Source sample: {result.source_sample_file.name}"
        )
        self._append_log(
            f"Source sheet: {result.source_sheet}"
        )
        self._append_log(
            f"Destination headers: {len(result.destination_headers)}"
        )
        self._append_log(
            f"Source headers: {len(result.source_headers)}"
        )
        self._append_log(
            f"Common headers available: {len(result.common_headers)}"
        )
        self._append_log(
            "Select stable identifier fields under Match Fields."
        )
        self._append_log(
            "Select fields that may differ under Write Fields."
        )
        self._append_log(
            "Do not select the same field in both lists."
        )
        self._append_log(
            "Ready to synchronize."
        )

    def on_analysis_failed(
        self,
        message: str,
    ):
        self.analyze_button.setEnabled(
            True
        )
        self.start_button.setEnabled(
            False
        )

        self._append_log(
            f"Analysis error: {message}"
        )

        QMessageBox.warning(
            self,
            "Header Analysis Failed",
            str(message),
        )

    @staticmethod
    def _selected_items(
        widget: QListWidget,
    ) -> tuple[str, ...]:
        return tuple(
            item.text()
            for item
            in widget.selectedItems()
        )

    def start_synchronization(self):
        if not self._paths_valid():
            return

        if self.analysis is None:
            QMessageBox.information(
                self,
                "Analyze Headers First",
                "Please run Analyze Headers before synchronization.",
            )
            return

        match_headers = self._selected_items(
            self.match_list
        )

        write_headers = self._selected_items(
            self.write_list
        )

        if not match_headers:
            QMessageBox.warning(
                self,
                "Missing Match Fields",
                "Select at least one Match field.",
            )
            return

        if not write_headers:
            QMessageBox.warning(
                self,
                "Missing Write Fields",
                "Select at least one Write field.",
            )
            return

        overlap = sorted(
            set(match_headers)
            & set(write_headers)
        )

        if overlap:
            QMessageBox.warning(
                self,
                "Field Selection Conflict",
                (
                    "The same field cannot be used as both a Match field "
                    "and a Write field.\n\nRemove these fields from one side:\n- "
                    + "\n- ".join(overlap)
                ),
            )
            return

        destination_sheet = self._selected_destination_sheet_name()
        source_sheet = self._selected_source_sheet_name()

        if not destination_sheet or not source_sheet:
            QMessageBox.warning(
                self,
                "Worksheet Selection Required",
                "Select both a Destination worksheet and a Source worksheet.",
            )
            return

        try:
            protection, worksheet_password, accepted = (
                self._get_destination_protection_password()
            )
        except Exception as exc:
            QMessageBox.warning(
                self,
                "Worksheet Protection Check Failed",
                str(exc),
            )
            return

        if not accepted:
            return

        self.progress.setValue(0)
        self.log.clear()

        self.analyze_button.setEnabled(
            False
        )
        self.start_button.setEnabled(
            False
        )

        self._append_log(
            "Task started"
        )
        self._append_log(
            "Match fields: "
            + " + ".join(match_headers)
        )
        self._append_log(
            "Write fields: "
            + ", ".join(write_headers)
        )
        self._append_log(
            "Append unmatched source rows: "
            + str(
                self.append_unmatched.isChecked()
            )
        )
        self._append_log(
            "Ignore blank source updates: "
            + str(
                self.ignore_blank_updates.isChecked()
            )
        )
        self._append_log(
            f"Destination worksheet: {destination_sheet}"
        )
        self._append_log(
            f"Source worksheet: {source_sheet}"
        )
        self._append_log(
            "Destination worksheet protection: "
            + (
                "Protected"
                if protection.get("is_protected", False)
                else "Not protected"
            )
        )
        self._append_log(
            "Important: fields expected to change belong in Write Fields, "
            "not Match Fields."
        )

        self.task_manager.start(
            lambda context: self.service.run(
                destination_file=self.destination_edit.text().strip(),
                source_folder=self.source_edit.text().strip(),
                destination_sheet=destination_sheet,
                source_sheet=source_sheet,
                match_headers=match_headers,
                write_headers=write_headers,
                append_unmatched=self.append_unmatched.isChecked(),
                ignore_blank_updates=self.ignore_blank_updates.isChecked(),
                worksheet_password=worksheet_password,
                context=context,
            ),
            {
                "progress": self.on_progress,
                "log": self._append_log,
                "completed": self.on_complete,
                "failed": self.on_failed,
            },
        )

    def on_progress(
        self,
        percent: int,
        message: str,
    ):
        self.progress.setValue(
            int(percent)
        )

    def on_complete(
        self,
        result,
    ):
        self.progress.setValue(
            100
        )

        self.analyze_button.setEnabled(
            True
        )
        self.start_button.setEnabled(
            True
        )

        self._append_log("")
        self._append_log(
            "COMPLETED"
        )
        self._append_log(
            result.message
        )

        output_file = (
            result.output_files[0]
            if getattr(
                result,
                "output_files",
                None,
            )
            else ""
        )

        QMessageBox.information(
            self,
            "Synchronization Completed",
            (
                "Updated workbook created successfully."
                + (
                    f"\n\n{output_file}"
                    if output_file
                    else ""
                )
            ),
        )

    def on_failed(
        self,
        message: str,
    ):
        self.analyze_button.setEnabled(
            True
        )
        self.start_button.setEnabled(
            True
        )

        self._append_log("")
        self._append_log(
            f"ERROR: {message}"
        )

        QMessageBox.warning(
            self,
            "Synchronization Failed",
            str(message),
        )
