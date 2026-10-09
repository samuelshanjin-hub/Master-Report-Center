from __future__ import annotations

import os
from pathlib import Path

from PyQt6.QtCore import Qt, QRect
from PyQt6.QtGui import QColor, QPainter

from PyQt6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QInputDialog,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QTextEdit,
    QVBoxLayout,
    QBoxLayout,
    QWidget,
)

from .service import ConsolidateReportsService


class DualColorProgressBar(QProgressBar):
    """
    Progress text stays readable over both the white/unfilled area and the
    blue/filled area: black on the unfilled portion, white on the filled portion.
    """

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
            ratio = (value - minimum) / (maximum - minimum)
            ratio = max(0.0, min(1.0, ratio))

        text = f"{value}%"
        rect = self.rect()

        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.TextAntialiasing, True)

        # Black text over the unfilled/white portion.
        painter.setPen(QColor("#17231D"))
        painter.drawText(rect, int(Qt.AlignmentFlag.AlignCenter), text)

        # White text over the filled/blue portion.
        filled_width = int(rect.width() * ratio)
        if filled_width > 0:
            painter.save()
            painter.setClipRect(QRect(0, 0, filled_width, rect.height()))
            painter.setPen(QColor("#FFFFFF"))
            painter.drawText(rect, int(Qt.AlignmentFlag.AlignCenter), text)
            painter.restore()

        painter.end()


class ConsolidateReportsPage(QWidget):
    """Publish-style Consolidate Reports page for Master Report Center Core V2."""

    def __init__(self, info, task_manager, logger, back_callback):
        super().__init__()

        self.info = info
        self.task_manager = task_manager
        self.logger = logger
        self.back_callback = back_callback
        self.service = ConsolidateReportsService()
        self._worksheet_password_cache = None

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
        root.setContentsMargins(22, 18, 22, 18)
        root.setSpacing(10)

        top = QHBoxLayout()

        back = QPushButton("← Report Center")
        back.setObjectName("Secondary")
        back.clicked.connect(self.back_callback)

        title_box = QVBoxLayout()
        title = QLabel("Consolidate Reports")
        title.setObjectName("PageTitle")

        subtitle = QLabel(
            "Combine multiple reports with matching headers into one consolidated report."
        )
        subtitle.setObjectName("Muted")

        title_box.addWidget(title)
        title_box.addWidget(subtitle)

        top.addWidget(back)
        top.addSpacing(10)
        top.addLayout(title_box)
        top.addStretch(1)

        root.addLayout(top)

        self.progress = DualColorProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        root.addWidget(self.progress)

        self.content_layout = QHBoxLayout()
        self.content_layout.setSpacing(12)

        self.left_layout = QVBoxLayout()
        self.left_layout.setSpacing(10)

        # 1. Source
        self.source_group = QGroupBox("1. Source")
        source_layout = QVBoxLayout(self.source_group)

        source_row = QHBoxLayout()
        self.source_edit = QLineEdit()
        self.source_edit.setPlaceholderText("Select folder containing source reports")

        source_browse = QPushButton("Browse")
        source_browse.setObjectName("Secondary")
        source_browse.clicked.connect(self.select_source_folder)

        source_open = QPushButton("Open Folder")
        source_open.setObjectName("Info")
        source_open.clicked.connect(
            lambda: self.open_folder(self.source_edit.text().strip())
        )

        source_row.addWidget(self.source_edit, 1)
        source_row.addWidget(source_browse)
        source_row.addWidget(source_open)

        source_layout.addLayout(source_row)

        source_note = QLabel(
            "All supported Excel reports in this folder will be analyzed."
        )
        source_note.setObjectName("Muted")
        source_layout.addWidget(source_note)

        self.left_layout.addWidget(self.source_group)

        # 2. Output
        self.output_group = QGroupBox("2. Output")
        output_layout = QVBoxLayout(self.output_group)

        output_row = QHBoxLayout()
        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("Select output folder")

        output_browse = QPushButton("Browse")
        output_browse.setObjectName("Secondary")
        output_browse.clicked.connect(self.select_output_folder)

        output_open = QPushButton("Open Folder")
        output_open.setObjectName("Info")
        output_open.clicked.connect(
            lambda: self.open_folder(self.output_edit.text().strip())
        )

        output_row.addWidget(self.output_edit, 1)
        output_row.addWidget(output_browse)
        output_row.addWidget(output_open)

        output_layout.addLayout(output_row)

        output_note = QLabel(
            "The first valid source workbook is copied and used as the output format template."
        )
        output_note.setWordWrap(True)
        output_note.setObjectName("Muted")
        output_layout.addWidget(output_note)

        self.left_layout.addWidget(self.output_group)

        # 3. Worksheet
        self.sheet_group = QGroupBox("3. Worksheet")
        sheet_layout = QVBoxLayout(self.sheet_group)

        sheet_label = QLabel("Worksheet name (optional)")
        self.worksheet_edit = QLineEdit()
        self.worksheet_edit.setPlaceholderText(
            "Leave blank to use the first worksheet"
        )
        self.worksheet_edit.textChanged.connect(self._clear_password_cache)

        sheet_layout.addWidget(sheet_label)
        sheet_layout.addWidget(self.worksheet_edit)

        sheet_note = QLabel(
            "If entered, the same worksheet name must exist in each source workbook."
        )
        sheet_note.setWordWrap(True)
        sheet_note.setObjectName("Muted")
        sheet_layout.addWidget(sheet_note)

        self.left_layout.addWidget(self.sheet_group)

        # 4. Data Handling
        self.data_group = QGroupBox("4. Data Handling")
        data_layout = QVBoxLayout(self.data_group)

        self.ignore_blank_rows = QCheckBox(
            "Ignore completely blank rows"
        )
        self.ignore_blank_rows.setChecked(True)

        self.remove_duplicates = QCheckBox(
            "Remove exact duplicate rows"
        )
        self.remove_duplicates.setChecked(True)

        self.add_source_file = QCheckBox(
            "Add source filename column"
        )
        self.add_source_file.setChecked(False)

        data_layout.addWidget(self.ignore_blank_rows)
        data_layout.addWidget(self.remove_duplicates)
        data_layout.addWidget(self.add_source_file)

        data_note = QLabel(
            "Header order may differ between files, but the same normalized header set is required."
        )
        data_note.setWordWrap(True)
        data_note.setObjectName("Muted")
        data_layout.addWidget(data_note)

        self.left_layout.addWidget(self.data_group)

        # Start Consolidation
        self.action_group = QGroupBox("Start Consolidation")
        self.action_layout = QBoxLayout(QBoxLayout.Direction.LeftToRight, self.action_group)
        action_layout = self.action_layout

        self.analyze_button = QPushButton("Analyze Reports")
        self.analyze_button.setObjectName("Info")
        self.analyze_button.clicked.connect(self.analyze)

        self.start_button = QPushButton("▶  Start Consolidation")
        self.start_button.setObjectName("Primary")
        self.start_button.clicked.connect(self.start_consolidation)

        action_layout.addWidget(self.analyze_button)
        action_layout.addWidget(self.start_button)

        self.left_layout.addWidget(self.action_group)
        self.left_layout.addStretch(1)

        # Activity log
        self.log_group = QGroupBox("Activity Log")
        log_layout = QVBoxLayout(self.log_group)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(320)
        self.log.setPlaceholderText(
            "Analysis and consolidation activity will appear here."
        )

        log_actions = QHBoxLayout()
        log_actions.addStretch(1)

        self.clear_log_button = QPushButton("Clear")
        self.clear_log_button.setObjectName("Secondary")
        self.clear_log_button.clicked.connect(self.log.clear)
        log_actions.addWidget(self.clear_log_button)

        log_layout.addLayout(log_actions)
        log_layout.addWidget(self.log)

        self.content_layout.addLayout(self.left_layout, 3)
        self.content_layout.addWidget(self.log_group, 2)

        root.addLayout(self.content_layout, 1)

        self._compact_layout = None
        self._apply_responsive_layout(force=True)

    def _apply_responsive_layout(self, force: bool = False):
        if not hasattr(self, "page_scroll"):
            return

        width = self.page_scroll.viewport().width()
        compact = width < 1180

        if not force and compact == self._compact_layout:
            return

        if compact:
            self.content_layout.setDirection(QBoxLayout.Direction.TopToBottom)
            self.content_layout.setStretch(0, 0)
            self.content_layout.setStretch(1, 0)

            self.action_layout.setDirection(QBoxLayout.Direction.TopToBottom)
            self.log.setMinimumHeight(360)
        else:
            self.content_layout.setDirection(QBoxLayout.Direction.LeftToRight)
            self.content_layout.setStretch(0, 3)
            self.content_layout.setStretch(1, 2)

            self.action_layout.setDirection(QBoxLayout.Direction.LeftToRight)
            self.log.setMinimumHeight(320)

        self._compact_layout = compact

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._apply_responsive_layout()

    def _append_log(self, message: str):
        self.log.append(str(message))

    def select_source_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Source Reports Folder",
            self.source_edit.text().strip() or str(Path.home()),
        )
        if folder:
            self.source_edit.setText(folder)
            self._worksheet_password_cache = None
            if not self.output_edit.text().strip():
                self.output_edit.setText(folder)
            self._append_log(f"Source folder selected: {folder}")

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Output Folder",
            self.output_edit.text().strip()
            or self.source_edit.text().strip()
            or str(Path.home()),
        )
        if folder:
            self.output_edit.setText(folder)
            self._append_log(f"Output folder selected: {folder}")

    def open_folder(self, folder: str):
        folder = folder.strip()
        if not folder:
            return

        path = Path(folder)
        if not path.exists():
            QMessageBox.warning(
                self,
                "Folder not found",
                str(path),
            )
            return

        try:
            os.startfile(str(path))
        except Exception as exc:
            QMessageBox.warning(
                self,
                "Unable to open folder",
                str(exc),
            )

    def _clear_password_cache(self):
        self._worksheet_password_cache = None

    def _get_template_protection_password(
        self,
        source_folder: str,
        worksheet_name: str,
    ):
        protection = self.service.get_template_protection_info(
            source_folder=source_folder,
            worksheet_name=worksheet_name,
        )

        if not protection.get("is_protected", False):
            return protection, "", True

        if self._worksheet_password_cache is not None:
            return protection, self._worksheet_password_cache, True

        if self.service.validate_template_worksheet_password(
            source_file=protection["template_file"],
            sheet_name=protection["sheet_name"],
            password="",
        ):
            self._worksheet_password_cache = ""
            return protection, "", True

        password, accepted = QInputDialog.getText(
            self,
            "Worksheet Protection",
            (
                "The output template worksheet is protected.\n\n"
                "A worksheet password is required to consolidate the report.\n"
                "The password is used only during this program session "
                "and is never saved or logged."
            ),
            QLineEdit.EchoMode.Password,
        )

        if not accepted:
            return protection, "", False

        if not self.service.validate_template_worksheet_password(
            source_file=protection["template_file"],
            sheet_name=protection["sheet_name"],
            password=password,
        ):
            QMessageBox.warning(
                self,
                "Incorrect Password",
                (
                    "The worksheet could not be unlocked with the password entered.\n\n"
                    "No source workbook changes were made."
                ),
            )
            return protection, "", False

        self._worksheet_password_cache = password
        return protection, password, True

    def analyze(self):
        source = self.source_edit.text().strip()

        if not source:
            QMessageBox.warning(
                self,
                "Missing Source Folder",
                "Please select the source reports folder.",
            )
            return

        try:
            result = self.service.analyze(source)

            self.log.clear()
            self._append_log("========== Analysis ==========")
            self._append_log(f"Source folder: {source}")
            self._append_log(
                f"Excel files found: {result['file_count']}"
            )

            for filename in result["files"]:
                self._append_log(f"  - {filename}")

            if result["file_count"] == 0:
                self._append_log("")
                self._append_log(
                    "No supported Excel files were found."
                )
            else:
                self._append_log("")
                self._append_log(
                    "Ready to start Excel COM consolidation."
                )

        except Exception as exc:
            self._append_log(f"Analysis error: {exc}")
            QMessageBox.warning(
                self,
                "Analysis Failed",
                str(exc),
            )

    def start_consolidation(self):
        source = self.source_edit.text().strip()
        output = self.output_edit.text().strip()
        worksheet = self.worksheet_edit.text().strip()

        if not source:
            QMessageBox.warning(
                self,
                "Missing Source Folder",
                "Please select the source reports folder.",
            )
            return

        if not output:
            QMessageBox.warning(
                self,
                "Missing Output Folder",
                "Please select the output folder.",
            )
            return

        try:
            protection, worksheet_password, accepted = (
                self._get_template_protection_password(
                    source_folder=source,
                    worksheet_name=worksheet,
                )
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
        self.start_button.setEnabled(False)
        self.analyze_button.setEnabled(False)

        self._append_log("Task started")
        self._append_log(f"Source folder: {source}")
        self._append_log(f"Output folder: {output}")
        self._append_log(
            f"Worksheet: {worksheet or '(First worksheet)'}"
        )
        self._append_log(
            "Ignore completely blank rows: "
            + str(self.ignore_blank_rows.isChecked())
        )
        self._append_log(
            "Remove exact duplicate rows: "
            + str(self.remove_duplicates.isChecked())
        )
        self._append_log(
            "Add source filename column: "
            + str(self.add_source_file.isChecked())
        )
        self._append_log(
            "Worksheet protection: "
            + ("Protected" if protection.get("is_protected", False) else "Not protected")
        )
        self._append_log("Microsoft Excel will run in the background.")

        self.task_manager.start(
            lambda context: self.service.run(
                source_folder=source,
                output_folder=output,
                worksheet_name=worksheet,
                remove_exact_duplicates=self.remove_duplicates.isChecked(),
                ignore_blank_rows=self.ignore_blank_rows.isChecked(),
                add_source_filename=self.add_source_file.isChecked(),
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

    def on_progress(self, percent: int, message: str):
        self.progress.setValue(int(percent))

    def on_complete(self, result):
        self.progress.setValue(100)
        self.start_button.setEnabled(True)
        self.analyze_button.setEnabled(True)

        self._append_log("")
        self._append_log("COMPLETED")
        self._append_log(result.message)

        output_text = ""
        if getattr(result, "output_files", None):
            output_text = result.output_files[0]

        QMessageBox.information(
            self,
            "Consolidation Completed",
            (
                "Combined workbook created successfully."
                + (
                    f"\n\n{output_text}"
                    if output_text
                    else ""
                )
            ),
        )

    def on_failed(self, message: str):
        self.start_button.setEnabled(True)
        self.analyze_button.setEnabled(True)

        self._append_log("")
        self._append_log(f"ERROR: {message}")

        QMessageBox.warning(
            self,
            "Consolidation Failed",
            str(message),
        )
