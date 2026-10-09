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
    LookupDataService,
    LookupHeaderAnalysis,
)


class DualColorProgressBar(QProgressBar):
    """Dual-colour percentage display matching Split / Consolidate / Sync."""

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


class LookupDataPage(QWidget):
    """
    Lookup / enrichment page specialized for:
        Reference Database -> Target workbook.
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

        self.service = LookupDataService()
        self.analysis: LookupHeaderAnalysis | None = None

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
            "Lookup Data"
        )
        title.setObjectName(
            "PageTitle"
        )

        top.addWidget(back)
        top.addWidget(title)
        top.addStretch(1)

        root.addLayout(top)

        subtitle = QLabel(
            "Enrich a Target workbook by looking up selected fields "
            "from a larger Reference Database."
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

        # --------------------------------------------------------------
        # 1. Reference Database
        # --------------------------------------------------------------
        self.reference_group = QGroupBox(
            "1. Reference Database"
        )
        reference_layout = QGridLayout(
            self.reference_group
        )

        self.reference_edit = QLineEdit()
        self.reference_edit.setPlaceholderText(
            "Select the larger Reference Database workbook"
        )

        reference_browse = QPushButton(
            "Browse..."
        )
        reference_browse.setObjectName(
            "Secondary"
        )
        reference_browse.clicked.connect(
            self.select_reference
        )

        reference_open = QPushButton(
            "Open Excel"
        )
        reference_open.setObjectName(
            "Green"
        )
        reference_open.clicked.connect(
            lambda: self.open_excel(
                self.reference_edit.text().strip()
            )
        )

        self.reference_sheet = QLineEdit()
        self.reference_sheet.setPlaceholderText(
            "Optional — blank uses first worksheet"
        )

        reference_layout.addWidget(
            QLabel("Reference Excel File"),
            0,
            0,
            1,
            3,
        )
        reference_layout.addWidget(
            self.reference_edit,
            1,
            0,
        )
        reference_layout.addWidget(
            reference_browse,
            1,
            1,
        )
        reference_layout.addWidget(
            reference_open,
            1,
            2,
        )
        reference_layout.addWidget(
            QLabel("Reference Worksheet"),
            2,
            0,
        )
        reference_layout.addWidget(
            self.reference_sheet,
            2,
            1,
            1,
            2,
        )

        # --------------------------------------------------------------
        # 2. Target
        # --------------------------------------------------------------
        self.target_group = QGroupBox(
            "2. Target"
        )
        target_layout = QGridLayout(
            self.target_group
        )

        self.target_edit = QLineEdit()
        self.target_edit.setPlaceholderText(
            "Select the smaller Target workbook to enrich"
        )

        target_browse = QPushButton(
            "Browse..."
        )
        target_browse.setObjectName(
            "Secondary"
        )
        target_browse.clicked.connect(
            self.select_target
        )

        target_open = QPushButton(
            "Open Excel"
        )
        target_open.setObjectName(
            "Green"
        )
        target_open.clicked.connect(
            lambda: self.open_excel(
                self.target_edit.text().strip()
            )
        )

        self.target_sheet = QLineEdit()
        self.target_sheet.setPlaceholderText(
            "Optional — blank uses first worksheet"
        )

        target_layout.addWidget(
            QLabel("Target Excel File"),
            0,
            0,
            1,
            3,
        )
        target_layout.addWidget(
            self.target_edit,
            1,
            0,
        )
        target_layout.addWidget(
            target_browse,
            1,
            1,
        )
        target_layout.addWidget(
            target_open,
            1,
            2,
        )
        target_layout.addWidget(
            QLabel("Target Worksheet"),
            2,
            0,
        )
        target_layout.addWidget(
            self.target_sheet,
            2,
            1,
            1,
            2,
        )

        # --------------------------------------------------------------
        # 3. Field Selection
        # --------------------------------------------------------------
        self.fields_group = QGroupBox(
            "3. Field Selection"
        )
        fields_layout = QGridLayout(
            self.fields_group
        )

        field_note = QLabel(
            "Analyze Headers first. Match Fields must exist in both files. "
            "Return Fields are values to retrieve from the Reference Database."
        )
        field_note.setWordWrap(True)
        field_note.setObjectName(
            "Muted"
        )

        self.match_list = QListWidget()
        self.match_list.setSelectionMode(
            QAbstractItemView.SelectionMode.ExtendedSelection
        )

        self.return_list = QListWidget()
        self.return_list.setSelectionMode(
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
            QLabel("Return Fields"),
            1,
            1,
        )
        fields_layout.addWidget(
            self.match_list,
            2,
            0,
        )
        fields_layout.addWidget(
            self.return_list,
            2,
            1,
        )

        # --------------------------------------------------------------
        # 4. Safe Write Behaviour
        # --------------------------------------------------------------
        self.options_group = QGroupBox(
            "4. Safe Write Behaviour"
        )
        options_layout = QVBoxLayout(
            self.options_group
        )

        self.fill_blank_only = QCheckBox(
            "Only fill blank Target cells"
        )
        self.fill_blank_only.setChecked(
            True
        )

        self.ignore_blank_reference = QCheckBox(
            "Ignore blank Reference values"
        )
        self.ignore_blank_reference.setChecked(
            True
        )

        options_layout.addWidget(
            self.fill_blank_only
        )
        options_layout.addWidget(
            self.ignore_blank_reference
        )

        safety_note = QLabel(
            "Lookup never appends unrelated Reference records to the Target. "
            "The original Target workbook is never overwritten."
        )
        safety_note.setWordWrap(True)
        safety_note.setObjectName(
            "Muted"
        )
        options_layout.addWidget(
            safety_note
        )

        # --------------------------------------------------------------
        # Start
        # --------------------------------------------------------------
        self.actions_group = QGroupBox(
            "Start Lookup"
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
            "▶ Start Lookup"
        )
        self.start_button.setObjectName(
            "Primary"
        )
        self.start_button.setEnabled(
            False
        )
        self.start_button.clicked.connect(
            self.start_lookup
        )

        actions_layout.addWidget(
            self.analyze_button
        )
        actions_layout.addWidget(
            self.start_button
        )

        self.grid.addWidget(
            self.reference_group,
            0,
            0,
        )
        self.grid.addWidget(
            self.target_group,
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

        # --------------------------------------------------------------
        # Activity Log — Clear at top-right, matching other modules.
        # --------------------------------------------------------------
        self.log_group = QGroupBox(
            "Activity Log"
        )
        log_layout = QVBoxLayout(
            self.log_group
        )

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(320)
        self.log.setPlaceholderText(
            "Header analysis and lookup activity will appear here."
        )

        log_actions = QHBoxLayout()
        log_actions.addStretch(1)

        clear = QPushButton(
            "Clear"
        )
        clear.setObjectName(
            "Secondary"
        )
        clear.clicked.connect(
            self.log.clear
        )
        log_actions.addWidget(clear)

        log_layout.addLayout(
            log_actions
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
            self.grid.addWidget(self.reference_group, 0, 0)
            self.grid.addWidget(self.target_group, 1, 0)
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
            self.grid.addWidget(self.reference_group, 0, 0)
            self.grid.addWidget(self.target_group, 0, 1)
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

    def _invalidate_analysis(self):
        self.analysis = None
        self.match_list.clear()
        self.return_list.clear()
        self.start_button.setEnabled(
            False
        )

    def select_reference(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Reference Database",
            self.reference_edit.text().strip()
            or str(Path.home()),
            "Excel Files (*.xlsx *.xlsm *.xlsb *.xls);;All Files (*.*)",
        )

        if path:
            self.reference_edit.setText(
                path
            )
            self._invalidate_analysis()
            self._append_log(
                f"Reference selected: {path}"
            )

    def select_target(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Target Workbook",
            self.target_edit.text().strip()
            or str(Path.home()),
            "Excel Files (*.xlsx *.xlsm *.xlsb *.xls);;All Files (*.*)",
        )

        if path:
            self.target_edit.setText(
                path
            )
            self._invalidate_analysis()
            self._append_log(
                f"Target selected: {path}"
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

    def _paths_valid(self) -> bool:
        if not self.reference_edit.text().strip():
            QMessageBox.warning(
                self,
                "Missing Reference Database",
                "Please select the Reference Database workbook.",
            )
            return False

        if not self.target_edit.text().strip():
            QMessageBox.warning(
                self,
                "Missing Target Workbook",
                "Please select the Target workbook.",
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
            "Analyzing Reference and Target headers..."
        )

        self.task_manager.start(
            lambda context: self.service.analyze(
                reference_file=self.reference_edit.text().strip(),
                target_file=self.target_edit.text().strip(),
                reference_sheet=self.reference_sheet.text().strip(),
                target_sheet=self.target_sheet.text().strip(),
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
        result: LookupHeaderAnalysis,
    ):
        self.analysis = result

        self.match_list.clear()
        self.return_list.clear()

        for header in result.common_headers:
            self.match_list.addItem(
                header
            )

        for header in result.return_headers:
            self.return_list.addItem(
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
            f"Reference sheet: {result.reference_sheet}"
        )
        self._append_log(
            f"Target sheet: {result.target_sheet}"
        )
        self._append_log(
            f"Reference headers: {len(result.reference_headers)}"
        )
        self._append_log(
            f"Target headers: {len(result.target_headers)}"
        )
        self._append_log(
            f"Common Match candidates: {len(result.common_headers)}"
        )
        self._append_log(
            "Select one or more stable Match Fields."
        )
        self._append_log(
            "Select the fields you want returned from the Reference Database."
        )
        self._append_log(
            "A Match field cannot also be selected as a Return field."
        )
        self._append_log(
            "Ready to run Lookup."
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
            for item in widget.selectedItems()
        )

    def start_lookup(self):
        if not self._paths_valid():
            return

        if self.analysis is None:
            QMessageBox.information(
                self,
                "Analyze Headers First",
                "Please run Analyze Headers before Lookup.",
            )
            return

        match_headers = self._selected_items(
            self.match_list
        )

        return_headers = self._selected_items(
            self.return_list
        )

        if not match_headers:
            QMessageBox.warning(
                self,
                "Missing Match Fields",
                "Select at least one Match field.",
            )
            return

        if not return_headers:
            QMessageBox.warning(
                self,
                "Missing Return Fields",
                "Select at least one Return field.",
            )
            return

        overlap = sorted(
            set(match_headers)
            & set(return_headers)
        )

        if overlap:
            QMessageBox.warning(
                self,
                "Field Selection Conflict",
                (
                    "A Match field cannot also be used as a Return field.\n\n"
                    "Remove these fields from one side:\n- "
                    + "\n- ".join(overlap)
                ),
            )
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
            "Return fields: "
            + ", ".join(return_headers)
        )
        self._append_log(
            "Only fill blank Target cells: "
            + str(
                self.fill_blank_only.isChecked()
            )
        )
        self._append_log(
            "Ignore blank Reference values: "
            + str(
                self.ignore_blank_reference.isChecked()
            )
        )

        self.task_manager.start(
            lambda context: self.service.run(
                reference_file=self.reference_edit.text().strip(),
                target_file=self.target_edit.text().strip(),
                reference_sheet=self.reference_sheet.text().strip(),
                target_sheet=self.target_sheet.text().strip(),
                match_headers=match_headers,
                return_headers=return_headers,
                fill_blank_only=self.fill_blank_only.isChecked(),
                ignore_blank_reference=self.ignore_blank_reference.isChecked(),
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
            "Lookup Completed",
            (
                "Lookup output created successfully."
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
            "Lookup Failed",
            str(message),
        )
