from __future__ import annotations

import os
from pathlib import Path

from PyQt6.QtCore import Qt, QRect
from PyQt6.QtGui import QColor, QPainter

from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QTextEdit,
    QVBoxLayout,
    QBoxLayout,
    QWidget,
)

from .service import SplitReportsService


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

        # Draw black text across the whole progress bar first.
        painter.setPen(QColor("#17231D"))
        painter.drawText(rect, int(Qt.AlignmentFlag.AlignCenter), text)

        # Redraw only the filled section in white.
        filled_width = int(rect.width() * ratio)
        if filled_width > 0:
            painter.save()
            painter.setClipRect(QRect(0, 0, filled_width, rect.height()))
            painter.setPen(QColor("#FFFFFF"))
            painter.drawText(rect, int(Qt.AlignmentFlag.AlignCenter), text)
            painter.restore()

        painter.end()


class SplitReportsPage(QWidget):
    """Original Publish Split Reports UI, connected to the Core V2 task framework."""

    def __init__(self, info, task_manager, logger, back_callback):
        super().__init__()
        self.info = info
        self.task_manager = task_manager
        self.logger = logger
        self.back_callback = back_callback
        self.file_path = ""
        self.preview_split_value = ""
        self.last_analysis = None
        self._worksheet_password_cache = None
        self.service = SplitReportsService()
        self._build()

    def _build(self):
        # Responsive shell:
        # large screens keep the familiar desktop arrangement;
        # smaller work areas switch to a vertical compact layout and scroll.
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
        title = QLabel("Split Reports")
        title.setObjectName("PageTitle")
        top.addWidget(back)
        top.addWidget(title)
        top.addStretch()
        root.addLayout(top)

        sub = QLabel("Split a large Excel report into multiple smaller reports based on a selected column.")
        sub.setObjectName("Muted")
        root.addWidget(sub)

        self.progress = DualColorProgressBar()
        self.progress.setValue(0)
        root.addWidget(self.progress)

        self.work = QHBoxLayout()
        self.work.setSpacing(12)

        self.config = QWidget()
        self.grid = QGridLayout(self.config)
        self.grid.setContentsMargins(0, 0, 0, 0)
        self.grid.setSpacing(10)

        self.source_group = QGroupBox("1. Source")
        sl = QGridLayout(self.source_group)
        self.file_edit = QLineEdit()
        self.file_edit.setPlaceholderText("Select an Excel workbook or paste its path")
        browse = QPushButton("Browse...")
        browse.setObjectName("Secondary")
        browse.clicked.connect(self.select_file)
        load = QPushButton("Load Location")
        load.setObjectName("Info")
        load.clicked.connect(self.load_path)
        open_btn = QPushButton("Open Excel")
        open_btn.setObjectName("Green")
        open_btn.clicked.connect(self.open_excel)
        sl.addWidget(QLabel("Source Excel File"), 0, 0, 1, 3)
        sl.addWidget(self.file_edit, 1, 0)
        sl.addWidget(browse, 1, 1)
        sl.addWidget(load, 1, 2)
        sl.addWidget(open_btn, 2, 2)

        self.output_group = QGroupBox("2. Output")
        ol = QGridLayout(self.output_group)
        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("Choose the parent folder for the split output")
        out_b = QPushButton("Browse...")
        out_b.setObjectName("Secondary")
        out_b.clicked.connect(self.select_output)
        ol.addWidget(QLabel("Output Folder"), 0, 0, 1, 2)
        ol.addWidget(self.output_edit, 1, 0)
        ol.addWidget(out_b, 1, 1)

        self.analysis_group = QGroupBox("3. Report Analysis")
        al = QGridLayout(self.analysis_group)
        self.sheet = QComboBox()
        self.sheet.currentTextChanged.connect(self._worksheet_changed)

        self.header_row = QComboBox()
        self.header_row.addItem("Auto Detect", None)
        for row_number in range(1, 51):
            self.header_row.addItem(f"Row {row_number}", row_number)
        self.header_row.currentIndexChanged.connect(self._header_row_changed)
        self._detected_header_row = 1

        self.header_status = QLabel("Detected: Row 1")
        self.header_status.setObjectName("Muted")

        self.column = QComboBox()
        self.column.currentTextChanged.connect(self.update_filename_preview)
        self.column.currentTextChanged.connect(self._invalidate_analysis)
        analyze_b = QPushButton("Analyze Report")
        analyze_b.setObjectName("Info")
        analyze_b.clicked.connect(self.analyze)
        al.addWidget(QLabel("Worksheet"), 0, 0)
        al.addWidget(self.sheet, 0, 1)
        al.addWidget(QLabel("Header Row"), 1, 0)
        al.addWidget(self.header_row, 1, 1)
        al.addWidget(self.header_status, 2, 1)
        al.addWidget(QLabel("Split Column"), 3, 0)
        al.addWidget(self.column, 3, 1)
        al.addWidget(analyze_b, 4, 1)

        self.options_group = QGroupBox("4. Options")
        opl = QVBoxLayout(self.options_group)
        self.exclude_hidden = QCheckBox("Exclude Hidden Columns")
        self.exclude_hidden.setChecked(True)

        self.include_other_sheets = QCheckBox("Include Other Worksheets")
        self.include_other_sheets.setChecked(True)
        self.include_other_sheets.setToolTip(
            "When disabled, all worksheets except the selected split worksheet "
            "will be removed from each output workbook. Cross-sheet formulas, "
            "Named Ranges or dropdown lists may be affected."
        )

        self.use_open = QCheckBox("Use Currently Open Excel Workbook View")
        self.use_open.stateChanged.connect(self._data_mode_changed)
        opl.addWidget(self.exclude_hidden)
        opl.addWidget(self.include_other_sheets)
        opl.addWidget(self.use_open)

        view_note = QLabel(
            "Current View uses only rows currently visible in the open Excel workbook "
            "after filtering or row hiding. Full Workbook mode ignores saved filters "
            "and reads all data rows."
        )
        view_note.setWordWrap(True)
        view_note.setObjectName("Muted")
        opl.addWidget(view_note)

        note = QLabel(
            "Original workbook formatting is preserved by the Excel reverse-filter "
            "engine. Header rows and any content above them are preserved. If "
            "dropdown Data Validation is detected, hidden supporting columns are "
            "kept automatically. Turning off Include Other Worksheets may affect "
            "cross-sheet formulas, Named Ranges or dropdown sources."
        )
        note.setWordWrap(True)
        note.setObjectName("Muted")
        opl.addWidget(note)

        self.grid.addWidget(self.source_group, 0, 0)
        self.grid.addWidget(self.output_group, 0, 1)
        self.grid.addWidget(self.analysis_group, 1, 0)
        self.grid.addWidget(self.options_group, 1, 1)

        self.actions_group = QGroupBox("Start Splitting")
        self.actions_layout = QGridLayout(self.actions_group)
        acl = self.actions_layout
        acl.setHorizontalSpacing(12)
        acl.setVerticalSpacing(8)

        self.actions_intro = QLabel("Review the analysis, then create one workbook per unique split value.")
        self.actions_intro.setObjectName("Muted")
        self.actions_intro.setWordWrap(True)
        acl.addWidget(self.actions_intro, 0, 0, 1, 4)

        self.suffix_label = QLabel("Custom filename suffix")
        self.filename_suffix = QLineEdit()
        self.filename_suffix.setPlaceholderText("Optional, e.g. Testing Outstanding")
        self.filename_suffix.textChanged.connect(self.update_filename_preview)

        self.add_date_suffix = QCheckBox("Add current date as final suffix")
        self.add_date_suffix.setChecked(True)
        self.add_date_suffix.stateChanged.connect(self.update_filename_preview)

        acl.addWidget(self.suffix_label, 1, 0)
        acl.addWidget(self.filename_suffix, 1, 1, 1, 2)
        acl.addWidget(self.add_date_suffix, 1, 3)

        self.preview_title = QLabel("Filename Preview")
        self.preview_title.setObjectName("Muted")
        self.filename_preview = QLabel("")
        self.filename_preview.setObjectName("FilenamePreview")
        self.filename_preview.setWordWrap(True)

        self.start = QPushButton("▶ Start Splitting")
        self.start.setObjectName("Primary")
        self.start.clicked.connect(self.start_split)

        acl.addWidget(self.preview_title, 2, 0)
        acl.addWidget(self.filename_preview, 2, 1, 1, 2)
        acl.addWidget(self.start, 2, 3)

        self.grid.addWidget(self.actions_group, 2, 0, 1, 2)
        self.update_filename_preview()

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(300)
        self.logbox = QGroupBox("Activity Log")
        ll = QVBoxLayout(self.logbox)
        clear = QPushButton("Clear")
        clear.setObjectName("Secondary")
        clear.clicked.connect(self.log.clear)
        ll.addWidget(clear, 0, Qt.AlignmentFlag.AlignRight)
        ll.addWidget(self.log)

        self.work.addWidget(self.config, 3)
        self.work.addWidget(self.logbox, 2)
        root.addLayout(self.work, 1)

        self._compact_layout = None
        self._apply_responsive_layout(force=True)

    def _clear_layout_items(self, layout):
        """Remove widgets/layout items without deleting the widgets."""
        while layout.count():
            layout.takeAt(0)

    def _set_actions_compact(self, compact: bool):
        """Reflow Start Splitting controls for narrow work areas."""
        self._clear_layout_items(self.actions_layout)

        if compact:
            self.actions_layout.addWidget(self.actions_intro, 0, 0, 1, 2)

            self.actions_layout.addWidget(self.suffix_label, 1, 0, 1, 2)
            self.actions_layout.addWidget(self.filename_suffix, 2, 0, 1, 2)

            self.actions_layout.addWidget(self.add_date_suffix, 3, 0, 1, 2)

            self.actions_layout.addWidget(self.preview_title, 4, 0, 1, 2)
            self.actions_layout.addWidget(self.filename_preview, 5, 0, 1, 2)

            self.actions_layout.addWidget(self.start, 6, 0, 1, 2)
        else:
            self.actions_layout.addWidget(self.actions_intro, 0, 0, 1, 4)

            self.actions_layout.addWidget(self.suffix_label, 1, 0)
            self.actions_layout.addWidget(self.filename_suffix, 1, 1, 1, 2)
            self.actions_layout.addWidget(self.add_date_suffix, 1, 3)

            self.actions_layout.addWidget(self.preview_title, 2, 0)
            self.actions_layout.addWidget(self.filename_preview, 2, 1, 1, 2)
            self.actions_layout.addWidget(self.start, 2, 3)

    def _apply_responsive_layout(self, force: bool = False):
        """
        Desktop:
            Source | Output
            Analysis | Options
            Start Splitting
            config | Activity Log

        Compact:
            Source
            Output
            Analysis
            Options
            Start Splitting
            Activity Log
        """
        if not hasattr(self, "page_scroll"):
            return

        width = self.page_scroll.viewport().width()

        # A fairly generous threshold is intentional: on 125%/150% Windows DPI,
        # a laptop may still report a seemingly wide pixel viewport while the
        # effective control space is much tighter.
        compact = width < 1380

        if not force and compact == self._compact_layout:
            return

        self._clear_layout_items(self.grid)

        if compact:
            self.grid.addWidget(self.source_group, 0, 0)
            self.grid.addWidget(self.output_group, 1, 0)
            self.grid.addWidget(self.analysis_group, 2, 0)
            self.grid.addWidget(self.options_group, 3, 0)
            self.grid.addWidget(self.actions_group, 4, 0)

            self.grid.setColumnStretch(0, 1)
            self.grid.setColumnStretch(1, 0)

            self.work.setDirection(QBoxLayout.Direction.TopToBottom)
            self.work.setStretch(0, 0)
            self.work.setStretch(1, 0)

            self.log.setMinimumHeight(360)
            self._set_actions_compact(True)
        else:
            self.grid.addWidget(self.source_group, 0, 0)
            self.grid.addWidget(self.output_group, 0, 1)
            self.grid.addWidget(self.analysis_group, 1, 0)
            self.grid.addWidget(self.options_group, 1, 1)
            self.grid.addWidget(self.actions_group, 2, 0, 1, 2)

            self.grid.setColumnStretch(0, 1)
            self.grid.setColumnStretch(1, 1)

            self.work.setDirection(QBoxLayout.Direction.LeftToRight)
            self.work.setStretch(0, 3)
            self.work.setStretch(1, 2)

            self.log.setMinimumHeight(300)
            self._set_actions_compact(False)

        self._compact_layout = compact

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._apply_responsive_layout()

    def _append_log(self, message: str):
        self.log.append(str(message))
        self.logger.info(str(message))

    def _invalidate_analysis(self):
        self.last_analysis = None

    def _clear_password_cache(self):
        self._worksheet_password_cache = None

    def _data_mode_changed(self):
        self._invalidate_analysis()
        self._clear_password_cache()

    def select_file(self):
        f, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            "",
            "Excel Files (*.xlsx *.xlsm *.xls)",
        )
        if f:
            self.set_file(f)

    def load_path(self):
        p = self.file_edit.text().strip().strip('"')
        if Path(p).is_dir():
            files = list(Path(p).glob("*.xlsx")) + list(Path(p).glob("*.xlsm")) + list(Path(p).glob("*.xls"))
            if len(files) == 1:
                p = str(files[0])
            elif len(files) > 1:
                f, _ = QFileDialog.getOpenFileName(
                    self,
                    "Select Excel File",
                    p,
                    "Excel Files (*.xlsx *.xlsm *.xls)",
                )
                p = f
        if p and Path(p).is_file():
            self.set_file(p)
        else:
            QMessageBox.warning(self, "Invalid path", "Please enter an existing Excel file or folder path.")

    def set_file(self, file_path: str):
        self.file_path = file_path
        self._invalidate_analysis()
        self._clear_password_cache()
        self.file_edit.setText(file_path)
        if not self.output_edit.text().strip():
            self.output_edit.setText(str(Path(file_path).parent))
        self.load_sheets()
        self._append_log(f"Selected: {file_path}")

    def load_sheets(self):
        try:
            names = self.service.get_sheet_names(self.file_path)
            self.sheet.blockSignals(True)
            self.sheet.clear()
            self.sheet.addItems(names)
            self.sheet.blockSignals(False)
            self.load_columns()
        except Exception as exc:
            QMessageBox.warning(self, "Unable to read workbook", str(exc))

    def _worksheet_changed(self):
        self._detected_header_row = 1
        self._invalidate_analysis()
        self._clear_password_cache()
        self.load_columns()

    def _header_row_changed(self):
        self._invalidate_analysis()
        self.load_columns()

    def _resolve_header_row(self) -> int:
        manual = self.header_row.currentData()
        if manual is not None:
            row_number = max(int(manual), 1)
            self.header_status.setText(f"Selected: Row {row_number}")
            return row_number

        sheet = self.sheet.currentText()
        if not self.file_path or not sheet:
            self._detected_header_row = 1
            self.header_status.setText("Detected: Row 1")
            return 1

        row_number = self.service.detect_header_row(
            self.file_path,
            sheet,
        )
        self._detected_header_row = max(int(row_number), 1)
        self.header_status.setText(
            f"Detected: Row {self._detected_header_row}"
        )
        return self._detected_header_row

    def _current_header_row(self) -> int:
        manual = self.header_row.currentData()
        if manual is not None:
            return max(int(manual), 1)
        return max(int(self._detected_header_row or 1), 1)

    def load_columns(self):
        try:
            sheet = self.sheet.currentText()
            if not sheet or not self.file_path:
                return
            header_row = self._resolve_header_row()
            columns = self.service.get_columns(
                self.file_path,
                sheet,
                header_row=header_row,
            )
            self._invalidate_analysis()
            self.column.clear()
            self.column.addItems(columns)
            # A useful default without hard-coding the engine itself.
            if columns:
                self.column.setCurrentIndex(0)
        except Exception as exc:
            self._append_log(f"Column load error: {exc}")

    def update_filename_preview(self):
        split_value = self.preview_split_value or "<Split Value>"
        suffix = ""

        if hasattr(self, "filename_suffix"):
            suffix = self.filename_suffix.text().strip()

        add_date = False
        if hasattr(self, "add_date_suffix"):
            add_date = self.add_date_suffix.isChecked()

        try:
            source_extension = Path(self.file_path).suffix if self.file_path else ".xlsx"
            filename = self.service.build_output_filename(
                split_value,
                custom_suffix=suffix,
                add_date_suffix=add_date,
                extension=source_extension,
            )
        except Exception:
            parts = [split_value]
            if suffix:
                parts.append(suffix)
            if add_date:
                from datetime import datetime
                parts.append(datetime.now().strftime("%d-%m-%Y"))
            source_extension = Path(self.file_path).suffix.lower() if self.file_path else ".xlsx"
            if source_extension not in {".xlsx", ".xlsm", ".xls"}:
                source_extension = ".xlsx"
            filename = " - ".join(parts) + source_extension

        if hasattr(self, "filename_preview"):
            self.filename_preview.setText(filename)

    def select_output(self):
        p = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if p:
            self.output_edit.setText(p)

    def open_excel(self):
        if not self.file_path:
            return
        try:
            os.startfile(self.file_path)
        except Exception as exc:
            QMessageBox.warning(self, "Open Excel", str(exc))

    def _get_protection_password(
        self,
        use_open_view: bool,
        purpose: str,
    ):
        """
        Return (protection_info, password, accepted).

        Password exists only in this page instance's memory and is never logged.
        Current View Analyze may require it because the VBA-equivalent engine
        must unprotect a temporary SaveCopyAs workbook before Excel can convert
        the current filtered view reliably.
        """
        protection = self.service.get_worksheet_protection_info(
            source_file=self.file_path,
            sheet_name=self.sheet.currentText(),
            use_open_excel_view=use_open_view,
        )

        if not protection["is_protected"]:
            return protection, "", True

        # Reuse a password already validated during this page session.
        if self._worksheet_password_cache is not None:
            return protection, self._worksheet_password_cache, True

        # Try blank password first, matching the verified VBA behaviour.
        if self.service.validate_worksheet_password(
            source_file=self.file_path,
            sheet_name=self.sheet.currentText(),
            password="",
            use_open_excel_view=use_open_view,
        ):
            self._worksheet_password_cache = ""
            return protection, "", True

        password, accepted = QInputDialog.getText(
            self,
            "Worksheet Protection",
            (
                "The selected worksheet is protected.\n\n"
                f"A worksheet password is required to {purpose}.\n"
                "The password is used only during this program session "
                "and is never saved or logged."
            ),
            QLineEdit.EchoMode.Password,
        )

        if not accepted:
            return protection, "", False

        if not self.service.validate_worksheet_password(
            source_file=self.file_path,
            sheet_name=self.sheet.currentText(),
            password=password,
            use_open_excel_view=use_open_view,
        ):
            QMessageBox.warning(
                self,
                "Incorrect Password",
                (
                    "The worksheet could not be unlocked with the password entered.\n\n"
                    "No source workbook changes or split files were created."
                ),
            )
            return protection, "", False

        self._worksheet_password_cache = password
        return protection, password, True

    def _block_protected_full_workbook_mode(self) -> bool:
        """
        Return True when processing must be stopped.

        FIX12 safety rule:
        Protected worksheets may only be processed through
        Current Excel Workbook View. Full Workbook mode is intentionally
        blocked to avoid treating saved filters as the active data set.
        """
        if self.use_open.isChecked():
            return False

        try:
            protection = self.service.get_worksheet_protection_info(
                source_file=self.file_path,
                sheet_name=self.sheet.currentText(),
                use_open_excel_view=False,
            )
        except Exception as exc:
            QMessageBox.warning(
                self,
                "Worksheet Protection Check Failed",
                str(exc),
            )
            return True

        if not protection.get("is_protected", False):
            return False

        QMessageBox.information(
            self,
            "Protected Worksheet",
            (
                "This worksheet is protected.\n\n"
                "Protected worksheets can only be processed using "
                "Current Excel Workbook View.\n\n"
                "Please open the workbook in Microsoft Excel, apply the "
                "required filters or row hiding, enable "
                "\"Use Currently Open Excel Workbook View\", then try again."
            ),
        )

        self._append_log("")
        self._append_log("Protected worksheet detected.")
        self._append_log(
            "Full Workbook Data mode is disabled for protected worksheets."
        )
        self._append_log(
            "Open the workbook in Excel and enable "
            "Use Currently Open Excel Workbook View."
        )

        return True

    def analyze(self):
        if not self.file_path or not self.sheet.currentText() or not self.column.currentText():
            QMessageBox.warning(
                self,
                "Missing information",
                "Select a source file, worksheet and split column first.",
            )
            return

        if self._block_protected_full_workbook_mode():
            return

        try:
            use_open_view = self.use_open.isChecked()
            worksheet_password = ""

            # Full Workbook analysis never needs to edit/unprotect anything.
            # Current View uses a VBA-equivalent SaveCopyAs conversion. If the
            # copied worksheet is protected, a password is therefore required
            # to convert that TEMPORARY copy only.
            if use_open_view:
                protection, worksheet_password, accepted = self._get_protection_password(
                    use_open_view=True,
                    purpose="analyze the current filtered Excel view",
                )
                if not accepted:
                    return
            else:
                protection = self.service.get_worksheet_protection_info(
                    source_file=self.file_path,
                    sheet_name=self.sheet.currentText(),
                    use_open_excel_view=False,
                )

            header_row = self._current_header_row()

            result = self.service.analyze(
                source_file=self.file_path,
                sheet_name=self.sheet.currentText(),
                split_column=self.column.currentText(),
                use_open_excel_view=use_open_view,
                worksheet_password=worksheet_password,
                header_row=header_row,
            )

            source_mode = result.get(
                "source_mode",
                "Current Excel Workbook View"
                if use_open_view
                else "Full Workbook Data",
            )

            self.last_analysis = {
                "file": str(Path(self.file_path).resolve()),
                "sheet": self.sheet.currentText(),
                "column": self.column.currentText(),
                "use_open_excel_view": use_open_view,
                "header_row": header_row,
                "rows": int(result["rows"]),
                "output_files": int(result["output_files"]),
                "source_mode": source_mode,
            }

            self._append_log("")
            self._append_log("========== Analysis ==========")
            self._append_log(f"Source mode: {source_mode}")
            self._append_log(f"Worksheet: {result['worksheet']}")
            self._append_log(f"Header row: {header_row}")
            self._append_log(f"Split column: {self.column.currentText()}")

            if use_open_view:
                self._append_log(f"Visible rows: {result['rows']:,}")
            else:
                self._append_log(f"Rows: {result['rows']:,}")

            self._append_log(f"Columns: {result['columns']:,}")
            self._append_log(f"Output files to create: {result['output_files']:,}")
            self._append_log(f"Blank split values: {result['blank_rows']:,}")
            self._append_log(
                "Worksheet protection: "
                + ("Protected" if protection["is_protected"] else "Not protected")
            )

            if use_open_view:
                self._append_log(
                    "Current View engine: Excel-native temporary view conversion"
                )

            self.preview_split_value = result.get("example_value", "") or ""
            self.update_filename_preview()
            self._append_log(f"Filename example: {self.filename_preview.text()}")
            self._append_log("Ready to split.")

        except Exception as exc:
            self.last_analysis = None
            self._append_log(f"Analysis error: {exc}")
            QMessageBox.warning(self, "Analysis failed", str(exc))


    def start_split(self):
        file_path = self.file_edit.text().strip().strip('"')
        if file_path and file_path != self.file_path and Path(file_path).is_file():
            self.file_path = file_path
            self._invalidate_analysis()
            self._clear_password_cache()

        if (
            not self.file_path
            or not self.output_edit.text().strip()
            or not self.sheet.currentText()
            or not self.column.currentText()
        ):
            QMessageBox.warning(
                self,
                "Missing information",
                "Select the source file, worksheet, split column and output folder.",
            )
            return

        if self._block_protected_full_workbook_mode():
            return

        use_open_view = self.use_open.isChecked()

        expected_analysis_rows = None
        expected_analysis_groups = None

        # Current View must be analyzed first because the data set is determined
        # by the live Excel filter/hidden-row state.
        if use_open_view:
            current_key = {
                "file": str(Path(self.file_path).resolve()),
                "sheet": self.sheet.currentText(),
                "column": self.column.currentText(),
                "header_row": self._current_header_row(),
                "use_open_excel_view": True,
            }

            saved_key = None
            if self.last_analysis:
                saved_key = {
                    "file": self.last_analysis.get("file"),
                    "sheet": self.last_analysis.get("sheet"),
                    "column": self.last_analysis.get("column"),
                    "header_row": self.last_analysis.get("header_row"),
                    "use_open_excel_view": self.last_analysis.get(
                        "use_open_excel_view"
                    ),
                }

            if saved_key != current_key:
                QMessageBox.information(
                    self,
                    "Analyze Current View First",
                    (
                        "Current Excel Workbook View is selected.\n\n"
                        "Please click Analyze Report first. The program will create "
                        "a temporary Excel-native view copy and verify the row/group "
                        "counts before splitting."
                    ),
                )
                return

            expected_analysis_rows = int(self.last_analysis["rows"])
            expected_analysis_groups = int(self.last_analysis["output_files"])

        elif self.last_analysis:
            current_key = {
                "file": str(Path(self.file_path).resolve()),
                "sheet": self.sheet.currentText(),
                "column": self.column.currentText(),
                "header_row": self._current_header_row(),
                "use_open_excel_view": False,
            }
            saved_key = {
                "file": self.last_analysis.get("file"),
                "sheet": self.last_analysis.get("sheet"),
                "column": self.last_analysis.get("column"),
                "header_row": self.last_analysis.get("header_row"),
                "use_open_excel_view": self.last_analysis.get(
                    "use_open_excel_view"
                ),
            }
            if saved_key == current_key:
                expected_analysis_rows = int(self.last_analysis["rows"])
                expected_analysis_groups = int(self.last_analysis["output_files"])

        if not self.include_other_sheets.isChecked():
            answer = QMessageBox.warning(
                self,
                "Other Worksheets Will Be Removed",
                (
                    "Include Other Worksheets is turned off.\n\n"
                    "Each output workbook will keep only the selected split "
                    "worksheet. Cross-sheet formulas, Named Ranges, charts or "
                    "dropdown lists that depend on other worksheets may no longer "
                    "work.\n\nContinue?"
                ),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return

        worksheet_password = ""

        try:
            protection, worksheet_password, accepted = self._get_protection_password(
                use_open_view=use_open_view,
                purpose="split the protected worksheet",
            )
            if not accepted:
                return
        except Exception as exc:
            QMessageBox.warning(
                self,
                "Worksheet Protection Check Failed",
                str(exc),
            )
            return

        opts = {
            "file": self.file_path,
            "sheet": self.sheet.currentText(),
            "column": self.column.currentText(),
            "base_output": self.output_edit.text().strip(),
            "exclude_hidden": self.exclude_hidden.isChecked(),
            "include_other_worksheets": self.include_other_sheets.isChecked(),
            "header_row": self._current_header_row(),
            "use_open_excel": use_open_view,
            "filename_suffix": self.filename_suffix.text().strip(),
            "add_date_suffix": self.add_date_suffix.isChecked(),
            "worksheet_password": worksheet_password,
            "expected_analysis_rows": expected_analysis_rows,
            "expected_analysis_groups": expected_analysis_groups,
        }

        self.start.setEnabled(False)
        self.progress.setValue(0)
        self.log.clear()
        self._append_log("Task started")
        self._append_log(
            "Data mode: "
            + (
                "Current Excel Workbook View"
                if use_open_view
                else "Full Workbook Data"
            )
        )
        self._append_log(f"Header row: {opts['header_row']}")
        self._append_log(
            "Other worksheets: "
            + ("Included" if opts["include_other_worksheets"] else "Removed")
        )
        self._append_log(
            "Worksheet protection: "
            + ("Protected" if protection["is_protected"] else "Not protected")
        )

        if use_open_view:
            self._append_log(
                "Current View engine: Excel-native temporary view conversion"
            )

        if expected_analysis_rows is not None:
            self._append_log(
                f"Analyze safety check: {expected_analysis_rows:,} row(s), "
                f"{expected_analysis_groups:,} output group(s)."
            )

        self.task_manager.start(
            lambda context: self.service.run(
                opts["file"],
                opts["sheet"],
                opts["column"],
                opts["base_output"],
                opts["exclude_hidden"],
                opts["use_open_excel"],
                opts["filename_suffix"],
                opts["add_date_suffix"],
                opts["worksheet_password"],
                context,
                expected_analysis_rows=opts["expected_analysis_rows"],
                expected_analysis_groups=opts["expected_analysis_groups"],
                header_row=opts["header_row"],
                include_other_worksheets=opts["include_other_worksheets"],
            ),
            {
                "progress": self.on_progress,
                "log": self._append_log,
                "completed": self.on_complete,
                "failed": self.on_failed,
            },
        )


    def on_progress(self, value: int, name: str):
        self.progress.setValue(value)
        if name:
            self._append_log(name)

    def on_complete(self, result):
        self.start.setEnabled(True)
        self.progress.setValue(100)
        self._append_log("")
        self._append_log("========== Completed ==========")
        self._append_log(result.message)
        output_folder = ""
        if result.output_files:
            output_folder = str(Path(result.output_files[0]).parent)
            self._append_log(f"Output folder: {output_folder}")
        try:
            if output_folder:
                os.startfile(output_folder)
        except Exception:
            pass
        QMessageBox.information(self, "Split Completed", f"{result.message}\n\n{output_folder}")

    def on_failed(self, message: str):
        self.start.setEnabled(True)
        self._append_log(f"ERROR: {message}")
        QMessageBox.critical(self, "Split failed", message)
