from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path
import gc
import shutil
import tempfile
import time
from typing import Any, Iterable

from core.models.job_result import JobResult


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb", ".xls"}
HEADER_ROW = 1

XL_PASTE_ALL = -4104
XL_PASTE_FORMATS = -4122
XL_UP = -4162
XL_TO_LEFT = -4159


def _normalise_header(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(
        str(value).replace("\r", " ").replace("\n", " ").split()
    ).casefold()


def _normalise_dedupe_value(value: Any) -> tuple[str, Any]:
    if value is None:
        return ("blank", "")
    if isinstance(value, str):
        if value == "":
            return ("blank", "")
        return ("text", value)
    if isinstance(value, bool):
        return ("bool", value)
    if isinstance(value, (int, float)):
        return ("number", float(value))
    return (type(value).__name__, str(value))


def _is_blank_row(row: Iterable[Any]) -> bool:
    for value in row:
        if value is None:
            continue
        if isinstance(value, str) and value == "":
            continue
        return False
    return True


def _to_matrix(value: Any, rows: int, cols: int) -> list[list[Any]]:
    if rows <= 0 or cols <= 0:
        return []
    if rows == 1 and cols == 1:
        if isinstance(value, (tuple, list)):
            if len(value) == 1 and isinstance(value[0], (tuple, list)):
                return [[value[0][0] if value[0] else None]]
            if len(value) == 1:
                return [[value[0]]]
        return [[value]]
    if rows == 1:
        if isinstance(value, (tuple, list)):
            if len(value) == 1 and isinstance(value[0], (tuple, list)):
                return [list(value[0])]
            return [list(value)]
        return [[value]]
    if cols == 1:
        if isinstance(value, (tuple, list)):
            result = []
            for item in value:
                if isinstance(item, (tuple, list)):
                    result.append([item[0] if item else None])
                else:
                    result.append([item])
            return result
        return [[value]]
    if not isinstance(value, (tuple, list)):
        return [[value]]
    return [
        list(row) if isinstance(row, (tuple, list)) else [row]
        for row in value
    ]


def _excel_files(folder: Path) -> list[Path]:
    return sorted(
        p for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in SUPPORTED_EXTENSIONS
        and not p.name.startswith("~$")
        and not p.name.startswith("Combined_Report_")
    )


_COM_BUSY_HRESULTS = {
    -2147418111,  # RPC_E_CALL_REJECTED: Call was rejected by callee.
    -2147417846,  # RPC_E_SERVERCALL_RETRYLATER: Server call retry later.
}


def _is_retryable_com_busy_error(exc: Exception) -> bool:
    hresult = getattr(exc, "hresult", None)
    if hresult in _COM_BUSY_HRESULTS:
        return True

    args = getattr(exc, "args", ()) or ()
    if args and isinstance(args[0], int) and args[0] in _COM_BUSY_HRESULTS:
        return True

    message = str(exc).casefold()
    return any(
        marker in message
        for marker in (
            "call was rejected by callee",
            "server call retry later",
            "application is busy",
            "message filter indicated that the application is busy",
        )
    )


def _call_with_com_retry(
    operation: str,
    func,
    log=None,
    attempts: int = 5,
    initial_delay: float = 0.5,
):
    """
    Retry only transient Excel COM busy-call failures.

    Non-busy COM errors are raised immediately.
    """
    attempts = max(int(attempts), 1)

    for attempt in range(1, attempts + 1):
        try:
            return func()
        except Exception as exc:
            if not _is_retryable_com_busy_error(exc) or attempt >= attempts:
                raise

            delay = initial_delay * attempt

            if log is not None:
                log(
                    f"Excel is busy during {operation}. "
                    f"Retrying COM operation ({attempt + 1}/{attempts}) "
                    f"in {delay:.1f}s..."
                )

            try:
                import pythoncom  # type: ignore
                pythoncom.PumpWaitingMessages()
            except Exception:
                pass

            time.sleep(delay)


def _safe_close(workbook: Any, save_changes: bool = False, log=None) -> None:
    if workbook is None:
        return
    try:
        _call_with_com_retry(
            "closing workbook",
            lambda: workbook.Close(SaveChanges=save_changes),
            log=log,
        )
    except Exception:
        pass


def _elapsed_text(seconds: float) -> str:
    return str(timedelta(seconds=seconds))


def _contiguous_mapping_segments(
    canonical_to_source: list[int],
) -> list[tuple[int, int, int, int]]:
    if not canonical_to_source:
        return []

    segments = []
    dest_start = dest_end = 1
    source_start = source_end = canonical_to_source[0]

    for dest_col in range(2, len(canonical_to_source) + 1):
        source_col = canonical_to_source[dest_col - 1]
        if source_col == source_end + 1:
            dest_end = dest_col
            source_end = source_col
        else:
            segments.append(
                (dest_start, dest_end, source_start, source_end)
            )
            dest_start = dest_end = dest_col
            source_start = source_end = source_col

    segments.append(
        (dest_start, dest_end, source_start, source_end)
    )
    return segments


class ConsolidateReportsService:
    """
    Project 003 Consolidate engine adapted to Master Report Center Core V2.

    V2:
    Python decides which rows to keep and how headers map.
    Excel COM performs native cell copy so source content and formatting are
    preserved as faithfully as Excel Copy/Paste supports.
    """

    @staticmethod
    def _safe_bool_property(obj, name: str, default: bool = False) -> bool:
        try:
            return bool(getattr(obj, name))
        except Exception:
            return default

    @staticmethod
    def _is_sheet_protected(ws) -> bool:
        values = []
        for attr in ("ProtectContents", "ProtectDrawingObjects", "ProtectScenarios"):
            try:
                values.append(bool(getattr(ws, attr)))
            except Exception:
                values.append(False)
        return any(values)

    def _capture_protection_settings(self, ws) -> dict:
        protected = self._is_sheet_protected(ws)
        state = {
            "is_protected": protected,
            "drawing_objects": self._safe_bool_property(ws, "ProtectDrawingObjects", False),
            "contents": self._safe_bool_property(ws, "ProtectContents", False),
            "scenarios": self._safe_bool_property(ws, "ProtectScenarios", False),
            "enable_selection": None,
            "allow_formatting_cells": False,
            "allow_formatting_columns": False,
            "allow_formatting_rows": False,
            "allow_inserting_columns": False,
            "allow_inserting_rows": False,
            "allow_inserting_hyperlinks": False,
            "allow_deleting_columns": False,
            "allow_deleting_rows": False,
            "allow_sorting": False,
            "allow_filtering": False,
            "allow_using_pivot_tables": False,
        }

        try:
            state["enable_selection"] = int(ws.EnableSelection)
        except Exception:
            pass

        if not protected:
            return state

        try:
            protection = ws.Protection
        except Exception:
            protection = None

        if protection is not None:
            mapping = {
                "allow_formatting_cells": "AllowFormattingCells",
                "allow_formatting_columns": "AllowFormattingColumns",
                "allow_formatting_rows": "AllowFormattingRows",
                "allow_inserting_columns": "AllowInsertingColumns",
                "allow_inserting_rows": "AllowInsertingRows",
                "allow_inserting_hyperlinks": "AllowInsertingHyperlinks",
                "allow_deleting_columns": "AllowDeletingColumns",
                "allow_deleting_rows": "AllowDeletingRows",
                "allow_sorting": "AllowSorting",
                "allow_filtering": "AllowFiltering",
                "allow_using_pivot_tables": "AllowUsingPivotTables",
            }
            for key, attr in mapping.items():
                state[key] = self._safe_bool_property(protection, attr, False)

        return state

    def _unprotect_sheet(self, ws, password: str, state: dict) -> None:
        if not state.get("is_protected", False):
            return

        _call_with_com_retry(
            "unlocking output template worksheet",
            lambda: ws.Unprotect(Password=password or ""),
        )

        if self._is_sheet_protected(ws):
            raise RuntimeError(
                "The output template worksheet is still protected after the unlock attempt."
            )

    def _restore_sheet_protection(self, ws, state: dict, password: str) -> None:
        if not state.get("is_protected", False):
            return

        try:
            ws.Protect(
                Password=password or "",
                DrawingObjects=bool(state.get("drawing_objects", False)),
                Contents=bool(state.get("contents", True)),
                Scenarios=bool(state.get("scenarios", False)),
                UserInterfaceOnly=False,
                AllowFormattingCells=bool(state.get("allow_formatting_cells", False)),
                AllowFormattingColumns=bool(state.get("allow_formatting_columns", False)),
                AllowFormattingRows=bool(state.get("allow_formatting_rows", False)),
                AllowInsertingColumns=bool(state.get("allow_inserting_columns", False)),
                AllowInsertingRows=bool(state.get("allow_inserting_rows", False)),
                AllowInsertingHyperlinks=bool(state.get("allow_inserting_hyperlinks", False)),
                AllowDeletingColumns=bool(state.get("allow_deleting_columns", False)),
                AllowDeletingRows=bool(state.get("allow_deleting_rows", False)),
                AllowSorting=bool(state.get("allow_sorting", False)),
                AllowFiltering=bool(state.get("allow_filtering", False)),
                AllowUsingPivotTables=bool(state.get("allow_using_pivot_tables", False)),
            )
        except Exception as exc:
            raise RuntimeError(
                "Worksheet protection could not be restored on the output workbook."
            ) from exc

        if not self._is_sheet_protected(ws):
            raise RuntimeError(
                "Worksheet protection could not be verified after restoration."
            )

        if state.get("enable_selection") is not None:
            try:
                ws.EnableSelection = int(state["enable_selection"])
            except Exception:
                pass

    def _read_valid_template_header(self, ws) -> tuple[list[str], list[str]]:
        last_col = int(
            ws.Cells(
                HEADER_ROW,
                ws.Columns.Count,
            ).End(XL_TO_LEFT).Column
        )

        header_raw = ws.Range(
            ws.Cells(HEADER_ROW, 1),
            ws.Cells(HEADER_ROW, last_col),
        ).Value2

        header_matrix = _to_matrix(header_raw, 1, last_col)
        source_headers = [
            "" if v is None else str(v)
            for v in header_matrix[0]
        ]
        source_normalised = [
            _normalise_header(v)
            for v in source_headers
        ]

        while source_normalised and source_normalised[-1] == "":
            source_normalised.pop()
            source_headers.pop()

        if not source_headers or any(h == "" for h in source_normalised):
            raise RuntimeError(
                "The header row contains one or more blank headers."
            )

        if len(set(source_normalised)) != len(source_normalised):
            raise RuntimeError(
                "The header row contains duplicate header names."
            )

        return source_headers, source_normalised

    def get_template_protection_info(
        self,
        source_folder: str,
        worksheet_name: str = "",
    ) -> dict:
        try:
            import pythoncom
            import win32com.client as win32
        except ImportError as exc:
            raise RuntimeError(
                "Consolidate Reports requires Microsoft Excel and pywin32 on Windows."
            ) from exc

        folder = Path(source_folder).expanduser().resolve()
        files = _excel_files(folder)
        if not files:
            raise RuntimeError(
                "No supported Excel files were found in the source folder."
            )

        pythoncom.CoInitialize()
        excel = wb = None
        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False

            last_error = None

            for file_path in files:
                try:
                    wb = _call_with_com_retry(
                        f"opening source workbook {file_path.name} for protection check",
                        lambda file_path=file_path: excel.Workbooks.Open(
                            str(file_path),
                            UpdateLinks=0,
                            ReadOnly=True,
                            AddToMru=False,
                            IgnoreReadOnlyRecommended=True,
                        ),
                    )

                    if worksheet_name.strip():
                        ws = wb.Worksheets(worksheet_name.strip())
                    else:
                        ws = wb.Worksheets(1)

                    self._read_valid_template_header(ws)
                    state = self._capture_protection_settings(ws)

                    return {
                        "is_protected": bool(state["is_protected"]),
                        "template_file": str(file_path),
                        "sheet_name": str(ws.Name),
                    }

                except Exception as exc:
                    last_error = exc

                finally:
                    _safe_close(wb, save_changes=False)
                    wb = None

            raise RuntimeError(
                "No valid template workbook could be identified."
                + (
                    f" Last error: {last_error}"
                    if last_error is not None
                    else ""
                )
            )

        finally:
            _safe_close(wb, save_changes=False)
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            gc.collect()
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def validate_template_worksheet_password(
        self,
        source_file: str,
        sheet_name: str,
        password: str,
    ) -> bool:
        try:
            import pythoncom
            import win32com.client as win32
        except ImportError as exc:
            raise RuntimeError(
                "Consolidate Reports requires Microsoft Excel and pywin32 on Windows."
            ) from exc

        source_path = Path(source_file).resolve()
        temp_dir = (
            Path(tempfile.gettempdir())
            / "GoldenFirReportCenter"
            / "ConsolidateReports"
            / "ProtectionCheck"
        )
        temp_dir.mkdir(parents=True, exist_ok=True)

        temp_file = temp_dir / (
            f"PROTECTION_CHECK_{datetime.now():%Y%m%d_%H%M%S_%f}_"
            f"{source_path.name}"
        )

        shutil.copy2(source_path, temp_file)

        pythoncom.CoInitialize()
        excel = wb = None

        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False

            wb = _call_with_com_retry(
                "opening temporary protection-check workbook",
                lambda: excel.Workbooks.Open(
                    str(temp_file),
                    UpdateLinks=0,
                    ReadOnly=False,
                    AddToMru=False,
                    IgnoreReadOnlyRecommended=True,
                ),
            )

            ws = wb.Worksheets(sheet_name)

            if not self._is_sheet_protected(ws):
                return True

            try:
                _call_with_com_retry(
                    "validating worksheet password",
                    lambda: ws.Unprotect(Password=password or ""),
                )
            except Exception:
                return False

            return not self._is_sheet_protected(ws)

        finally:
            _safe_close(wb, save_changes=False)
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            gc.collect()
            try:
                temp_file.unlink(missing_ok=True)
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def analyze(self, source_folder: str) -> dict:
        folder = Path(source_folder).expanduser()
        if not folder.exists():
            raise FileNotFoundError(f"Folder not found: {folder}")
        if not folder.is_dir():
            raise NotADirectoryError(f"Not a folder: {folder}")

        files = _excel_files(folder)
        return {
            "file_count": len(files),
            "files": [p.name for p in files],
        }

    @staticmethod
    def _write_persistent_log(
        log_file: Path,
        *,
        source_folder: Path,
        output_folder: Path,
        worksheet_name: str,
        remove_exact_duplicates: bool,
        ignore_blank_rows: bool,
        add_source_filename: bool,
        worksheet_protected: bool,
        files_found: int,
        files_processed: int,
        skipped_files: list[tuple[str, str]],
        source_rows_read: int,
        rows_written: int,
        duplicates_removed: int,
        blank_rows_skipped: int,
        output_file: Path | None,
        duration_seconds: float,
        status: str,
        error_message: str | None = None,
    ) -> None:
        lines = [
            "================================================",
            "Golden Fir Report Center - Consolidate Reports",
            "================================================",
            "",
            "Source Folder:",
            str(source_folder),
            "",
            "Output Folder:",
            str(output_folder),
            "",
            "Worksheet:",
            worksheet_name.strip() or "(First worksheet)",
            "",
            "Ignore Completely Blank Rows:",
            str(bool(ignore_blank_rows)),
            "",
            "Remove Exact Duplicate Rows:",
            str(bool(remove_exact_duplicates)),
            "",
            "Add Source Filename Column:",
            str(bool(add_source_filename)),
            "",
            "Template Worksheet Protected:",
            str(bool(worksheet_protected)),
            "",
            "Source Files Found:",
            str(files_found),
            "",
            "Source Files Processed:",
            str(files_processed),
            "",
            "Files Skipped:",
            str(len(skipped_files)),
            "",
        ]

        if skipped_files:
            lines.append("Skipped File Details:")
            for filename, reason in skipped_files:
                lines.append(f"- {filename} â€” {reason}")
            lines.append("")

        lines.extend([
            "Source Rows Read:",
            str(source_rows_read),
            "",
            "Rows Written:",
            str(rows_written),
            "",
            "Exact Duplicates Removed:",
            str(duplicates_removed),
            "",
            "Blank Rows Skipped:",
            str(blank_rows_skipped),
            "",
            "Output File:",
            str(output_file) if output_file is not None else "(None)",
            "",
            "Duration:",
            _elapsed_text(duration_seconds),
            "",
            "Status:",
            status,
        ])

        if error_message:
            lines.extend(["", "Error:", error_message])

        lines.extend(["", "================================================", ""])
        log_file.parent.mkdir(parents=True, exist_ok=True)
        log_file.write_text("\n".join(lines), encoding="utf-8")

    @staticmethod
    def _copy_source_row_fidelity(
        source_ws,
        source_row: int,
        output_ws,
        output_row: int,
        segments: list[tuple[int, int, int, int]],
    ) -> None:
        for (
            dest_start,
            dest_end,
            source_start,
            source_end,
        ) in segments:
            source_range = source_ws.Range(
                source_ws.Cells(source_row, source_start),
                source_ws.Cells(source_row, source_end),
            )
            destination_range = output_ws.Range(
                output_ws.Cells(output_row, dest_start),
                output_ws.Cells(output_row, dest_end),
            )
            _call_with_com_retry(
                "copying source row content",
                lambda: source_range.Copy(),
            )
            _call_with_com_retry(
                "pasting source row content",
                lambda: destination_range.PasteSpecial(XL_PASTE_ALL),
            )

        try:
            output_ws.Rows(output_row).RowHeight = (
                source_ws.Rows(source_row).RowHeight
            )
        except Exception:
            pass

        try:
            output_ws.Rows(output_row).Hidden = bool(
                source_ws.Rows(source_row).Hidden
            )
        except Exception:
            pass

    def run(
        self,
        source_folder: str,
        output_folder: str,
        worksheet_name: str,
        remove_exact_duplicates: bool,
        ignore_blank_rows: bool,
        add_source_filename: bool,
        worksheet_password: str = "",
        context=None,
    ) -> JobResult:
        try:
            import pythoncom
            import win32com.client as win32
        except ImportError as exc:
            raise RuntimeError(
                "Consolidate Reports requires Microsoft Excel and pywin32 on Windows."
            ) from exc

        started = time.perf_counter()

        source_folder_path = Path(source_folder).expanduser().resolve()
        output_folder_path = Path(output_folder).expanduser().resolve()
        output_folder_path.mkdir(parents=True, exist_ok=True)

        timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        log_file = output_folder_path / f"Consolidate_Log_{timestamp}.txt"

        files = []
        skipped_files: list[tuple[str, str]] = []

        source_rows_read = 0
        rows_written = 0
        duplicates_removed = 0
        blank_rows_skipped = 0
        files_processed = 0

        canonical_headers = None
        canonical_normalised = None
        template_file = None
        template_sheet_name = None
        template_last_row = HEADER_ROW
        template_protection_state = None
        output_file = None

        keep_plan: list[
            tuple[Path, str, int, list[int]]
        ] = []
        seen_rows = set()

        def report(percent: int, message: str) -> None:
            safe = max(0, min(100, int(percent)))
            context.progress(safe, message)
            context.log(message)

        if not source_folder_path.is_dir():
            raise NotADirectoryError(
                f"Source folder not found: {source_folder_path}"
            )

        files = _excel_files(source_folder_path)
        if not files:
            raise RuntimeError(
                "No supported Excel files were found in the source folder."
            )

        pythoncom.CoInitialize()
        excel = source_wb = output_wb = None

        try:
            report(1, "Starting Microsoft Excel...")

            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            excel.AskToUpdateLinks = False

            try:
                excel.Calculation = -4135
                report(2, "Excel calculation mode set to manual.")
            except Exception:
                report(
                    2,
                    "Excel calculation mode could not be changed; continuing normally.",
                )

            # Phase 1: analyze/plan.
            for file_index, file_path in enumerate(files, start=1):
                base_percent = 5 + int(
                    (file_index - 1) / len(files) * 35
                )
                report(base_percent, f"Analyzing {file_path.name}...")

                try:
                    source_wb = _call_with_com_retry(
                        f"opening source workbook {file_path.name}",
                        lambda file_path=file_path: excel.Workbooks.Open(
                            str(file_path),
                            UpdateLinks=0,
                            ReadOnly=True,
                            AddToMru=False,
                            IgnoreReadOnlyRecommended=True,
                        ),
                        log=context.log,
                    )

                    if worksheet_name.strip():
                        try:
                            ws = source_wb.Worksheets(
                                worksheet_name.strip()
                            )
                        except Exception as exc:
                            raise RuntimeError(
                                f"Worksheet '{worksheet_name.strip()}' was not found."
                            ) from exc
                    else:
                        ws = source_wb.Worksheets(1)

                    last_col = int(
                        ws.Cells(
                            HEADER_ROW,
                            ws.Columns.Count,
                        ).End(XL_TO_LEFT).Column
                    )

                    header_raw = ws.Range(
                        ws.Cells(HEADER_ROW, 1),
                        ws.Cells(HEADER_ROW, last_col),
                    ).Value2

                    header_matrix = _to_matrix(
                        header_raw,
                        1,
                        last_col,
                    )

                    source_headers = [
                        "" if v is None else str(v)
                        for v in header_matrix[0]
                    ]
                    source_normalised = [
                        _normalise_header(v)
                        for v in source_headers
                    ]

                    while (
                        source_normalised
                        and source_normalised[-1] == ""
                    ):
                        source_normalised.pop()
                        source_headers.pop()

                    if (
                        not source_headers
                        or any(h == "" for h in source_normalised)
                    ):
                        raise RuntimeError(
                            "The header row contains one or more blank headers."
                        )

                    if (
                        len(set(source_normalised))
                        != len(source_normalised)
                    ):
                        raise RuntimeError(
                            "The header row contains duplicate header names."
                        )

                    if canonical_headers is None:
                        canonical_headers = list(source_headers)
                        canonical_normalised = list(
                            source_normalised
                        )
                        template_file = file_path
                        template_sheet_name = str(ws.Name)
                        template_protection_state = self._capture_protection_settings(ws)

                        template_last_row = max(
                            HEADER_ROW,
                            int(
                                ws.Cells(
                                    ws.Rows.Count,
                                    1,
                                ).End(XL_UP).Row
                            ),
                        )

                        report(
                            base_percent,
                            f"Using {file_path.name} as the output format template.",
                        )
                    else:
                        if (
                            set(source_normalised)
                            != set(canonical_normalised)
                        ):
                            missing = [
                                canonical_headers[i]
                                for i, h in enumerate(
                                    canonical_normalised
                                )
                                if h not in source_normalised
                            ]
                            extra = [
                                source_headers[i]
                                for i, h in enumerate(
                                    source_normalised
                                )
                                if h not in canonical_normalised
                            ]
                            parts = []
                            if missing:
                                parts.append(
                                    "missing: " + ", ".join(missing)
                                )
                            if extra:
                                parts.append(
                                    "extra: " + ", ".join(extra)
                                )
                            raise RuntimeError(
                                "Header mismatch ("
                                + "; ".join(parts)
                                + ")"
                            )

                    source_index = {
                        h: i + 1
                        for i, h in enumerate(source_normalised)
                    }
                    canonical_to_source = [
                        source_index[h]
                        for h in canonical_normalised
                    ]

                    last_row = HEADER_ROW
                    for source_col in canonical_to_source:
                        candidate = int(
                            ws.Cells(
                                ws.Rows.Count,
                                source_col,
                            ).End(XL_UP).Row
                        )
                        last_row = max(
                            last_row,
                            candidate,
                        )

                    if last_row >= HEADER_ROW + 1:
                        source_col_count = len(source_headers)
                        data_raw = ws.Range(
                            ws.Cells(HEADER_ROW + 1, 1),
                            ws.Cells(
                                last_row,
                                source_col_count,
                            ),
                        ).Value2

                        matrix = _to_matrix(
                            data_raw,
                            last_row - HEADER_ROW,
                            source_col_count,
                        )
                    else:
                        matrix = []

                    for offset, source_row_values in enumerate(
                        matrix,
                        start=1,
                    ):
                        source_row_number = HEADER_ROW + offset
                        source_rows_read += 1

                        aligned = [
                            source_row_values[
                                source_col - 1
                            ]
                            for source_col
                            in canonical_to_source
                        ]

                        if (
                            ignore_blank_rows
                            and _is_blank_row(aligned)
                        ):
                            blank_rows_skipped += 1
                            continue

                        if remove_exact_duplicates:
                            dedupe_key = tuple(
                                _normalise_dedupe_value(v)
                                for v in aligned
                            )
                            if dedupe_key in seen_rows:
                                duplicates_removed += 1
                                continue
                            seen_rows.add(dedupe_key)

                        keep_plan.append(
                            (
                                file_path,
                                str(ws.Name),
                                source_row_number,
                                list(canonical_to_source),
                            )
                        )

                    files_processed += 1

                except Exception as exc:
                    reason = str(exc)
                    skipped_files.append(
                        (file_path.name, reason)
                    )
                    report(
                        base_percent,
                        f"Skipped {file_path.name}: {reason}",
                    )

                finally:
                    _safe_close(
                        source_wb,
                        save_changes=False,
                    )
                    source_wb = None

            if canonical_headers is None:
                raise RuntimeError(
                    "No source workbook with a valid header row was processed."
                )

            if (
                template_file is None
                or template_sheet_name is None
            ):
                raise RuntimeError(
                    "No valid template workbook was identified."
                )

            # Phase 2: create output.
            output_file = output_folder_path / (
                f"Combined_Report_Merged_{timestamp}"
                f"{template_file.suffix}"
            )

            report(
                42,
                f"Copying {template_file.name} as the output workbook template...",
            )
            shutil.copy2(
                template_file,
                output_file,
            )

            output_wb = _call_with_com_retry(
                f"opening output workbook {output_file.name}",
                lambda: excel.Workbooks.Open(
                    str(output_file),
                    UpdateLinks=0,
                    ReadOnly=False,
                    AddToMru=False,
                    IgnoreReadOnlyRecommended=True,
                ),
                log=context.log,
            )
            output_ws = output_wb.Worksheets(
                template_sheet_name
            )

            if template_protection_state is None:
                template_protection_state = self._capture_protection_settings(
                    output_ws
                )

            if template_protection_state.get("is_protected", False):
                report(
                    43,
                    "Output template worksheet is protected; unlocking output copy...",
                )
                self._unprotect_sheet(
                    output_ws,
                    worksheet_password,
                    template_protection_state,
                )
                context.log(
                    "Output template worksheet unlocked in the copied workbook only."
                )

            base_col_count = len(canonical_headers)
            final_col_count = (
                base_col_count + 1
                if add_source_filename
                else base_col_count
            )

            current_last_row = max(
                template_last_row,
                int(
                    output_ws.Cells(
                        output_ws.Rows.Count,
                        1,
                    ).End(XL_UP).Row
                ),
            )

            if current_last_row >= HEADER_ROW + 1:
                report(
                    44,
                    "Clearing template data area...",
                )
                _call_with_com_retry(
                    "clearing template data area",
                    lambda: output_ws.Range(
                        output_ws.Cells(
                            HEADER_ROW + 1,
                            1,
                        ),
                        output_ws.Cells(
                            current_last_row,
                            final_col_count,
                        ),
                    ).Clear(),
                    log=context.log,
                )

            if add_source_filename:
                _call_with_com_retry(
                    "copying Source File header format",
                    lambda: output_ws.Cells(
                        HEADER_ROW,
                        base_col_count,
                    ).Copy(),
                    log=context.log,
                )
                _call_with_com_retry(
                    "pasting Source File header format",
                    lambda: output_ws.Cells(
                        HEADER_ROW,
                        final_col_count,
                    ).PasteSpecial(XL_PASTE_FORMATS),
                    log=context.log,
                )
                _call_with_com_retry(
                    "writing Source File header",
                    lambda: setattr(
                        output_ws.Cells(HEADER_ROW, final_col_count),
                        "Value2",
                        "Source File",
                    ),
                    log=context.log,
                )

                try:
                    output_ws.Columns(
                        final_col_count
                    ).ColumnWidth = output_ws.Columns(
                        base_col_count
                    ).ColumnWidth
                except Exception:
                    pass

            by_file = {}
            for (
                file_path,
                source_sheet_name,
                source_row_number,
                column_map,
            ) in keep_plan:
                by_file.setdefault(
                    file_path,
                    [],
                ).append(
                    (
                        source_sheet_name,
                        source_row_number,
                        column_map,
                    )
                )

            total_rows_to_write = len(keep_plan)
            output_row = HEADER_ROW + 1
            processed_for_copy = 0

            for file_index, (
                file_path,
                rows_for_file,
            ) in enumerate(by_file.items(), start=1):
                report(
                    46 + int(
                        (file_index - 1)
                        / max(1, len(by_file))
                        * 48
                    ),
                    f"Copying Excel content and formatting from {file_path.name}...",
                )

                source_wb = _call_with_com_retry(
                    f"opening source workbook {file_path.name}",
                    lambda file_path=file_path: excel.Workbooks.Open(
                        str(file_path),
                        UpdateLinks=0,
                        ReadOnly=True,
                        AddToMru=False,
                        IgnoreReadOnlyRecommended=True,
                    ),
                    log=context.log,
                )

                try:
                    source_sheet_name = (
                        rows_for_file[0][0]
                    )
                    source_ws = source_wb.Worksheets(
                        source_sheet_name
                    )
                    column_map = rows_for_file[0][2]
                    segments = (
                        _contiguous_mapping_segments(
                            column_map
                        )
                    )

                    for (
                        _sheet_name,
                        source_row_number,
                        _column_map,
                    ) in rows_for_file:
                        self._copy_source_row_fidelity(
                            source_ws,
                            source_row_number,
                            output_ws,
                            output_row,
                            segments,
                        )

                        if add_source_filename:
                            target = output_ws.Cells(
                                output_row,
                                final_col_count,
                            )
                            try:
                                _call_with_com_retry(
                                    "copying source filename cell format",
                                    lambda: output_ws.Cells(
                                        output_row,
                                        base_col_count,
                                    ).Copy(),
                                    log=context.log,
                                )
                                _call_with_com_retry(
                                    "pasting source filename cell format",
                                    lambda: target.PasteSpecial(
                                        XL_PASTE_FORMATS
                                    ),
                                    log=context.log,
                                )
                            except Exception:
                                pass
                            _call_with_com_retry(
                                "writing source filename",
                                lambda: setattr(target, "Value2", file_path.name),
                                log=context.log,
                            )

                        output_row += 1
                        rows_written += 1
                        processed_for_copy += 1

                        if (
                            processed_for_copy == 1
                            or processed_for_copy % 25 == 0
                            or processed_for_copy
                            == total_rows_to_write
                        ):
                            report(
                                46 + int(
                                    processed_for_copy
                                    / max(
                                        1,
                                        total_rows_to_write,
                                    )
                                    * 48
                                ),
                                (
                                    f"Writing row "
                                    f"{processed_for_copy:,} "
                                    f"of {total_rows_to_write:,}..."
                                ),
                            )
                finally:
                    _safe_close(
                        source_wb,
                        save_changes=False,
                    )
                    source_wb = None

            try:
                excel.CutCopyMode = False
            except Exception:
                pass

            if (
                template_protection_state is not None
                and template_protection_state.get("is_protected", False)
            ):
                report(
                    95,
                    "Restoring worksheet protection on output workbook...",
                )
                self._restore_sheet_protection(
                    output_ws,
                    template_protection_state,
                    worksheet_password,
                )
                context.log(
                    "Worksheet protection restored and verified on output workbook."
                )

            report(
                96,
                f"Saving {output_file.name}...",
            )
            _call_with_com_retry(
                f"saving output workbook {output_file.name}",
                lambda: output_wb.Save(),
                log=context.log,
            )
            _safe_close(
                output_wb,
                save_changes=True,
                log=context.log,
            )
            output_wb = None

            elapsed = time.perf_counter() - started

            self._write_persistent_log(
                log_file,
                source_folder=source_folder_path,
                output_folder=output_folder_path,
                worksheet_name=worksheet_name,
                remove_exact_duplicates=remove_exact_duplicates,
                ignore_blank_rows=ignore_blank_rows,
                add_source_filename=add_source_filename,
                worksheet_protected=bool(
                    template_protection_state
                    and template_protection_state.get("is_protected", False)
                ),
                files_found=len(files),
                files_processed=files_processed,
                skipped_files=skipped_files,
                source_rows_read=source_rows_read,
                rows_written=rows_written,
                duplicates_removed=duplicates_removed,
                blank_rows_skipped=blank_rows_skipped,
                output_file=output_file,
                duration_seconds=elapsed,
                status="Completed Successfully",
            )

            context.log("")
            context.log(
                "========== Consolidation Summary =========="
            )
            context.log(
                f"Source files found: {len(files)}"
            )
            context.log(
                f"Source files processed: {files_processed}"
            )
            context.log(
                f"Files skipped: {len(skipped_files)}"
            )
            context.log(
                f"Source rows read: {source_rows_read}"
            )
            context.log(
                f"Rows written: {rows_written}"
            )
            context.log(
                f"Exact duplicates removed: "
                f"{duplicates_removed}"
            )
            context.log(
                f"Blank rows skipped: "
                f"{blank_rows_skipped}"
            )
            context.log(
                f"Duration: {_elapsed_text(elapsed)}"
            )
            context.log(
                f"Output workbook: {output_file}"
            )
            context.log(
                f"Persistent log: {log_file}"
            )

            report(
                100,
                "Consolidation completed.",
            )

            return JobResult(
                success=True,
                message=(
                    f"Consolidated "
                    f"{files_processed} file(s), "
                    f"{rows_written} row(s)."
                ),
                output_files=[
                    str(output_file),
                    str(log_file),
                ],
                processed_count=files_processed,
                skipped_count=len(skipped_files),
                warning_count=len(skipped_files),
            )

        except Exception as exc:
            elapsed = time.perf_counter() - started

            try:
                self._write_persistent_log(
                    log_file,
                    source_folder=source_folder_path,
                    output_folder=output_folder_path,
                    worksheet_name=worksheet_name,
                    remove_exact_duplicates=remove_exact_duplicates,
                    ignore_blank_rows=ignore_blank_rows,
                    add_source_filename=add_source_filename,
                    worksheet_protected=bool(
                        template_protection_state
                        and template_protection_state.get("is_protected", False)
                    ),
                    files_found=len(files),
                    files_processed=files_processed,
                    skipped_files=skipped_files,
                    source_rows_read=source_rows_read,
                    rows_written=rows_written,
                    duplicates_removed=duplicates_removed,
                    blank_rows_skipped=blank_rows_skipped,
                    output_file=output_file,
                    duration_seconds=elapsed,
                    status="Failed",
                    error_message=str(exc),
                )
                context.log(
                    f"Persistent failure log: "
                    f"{log_file}"
                )
            except Exception:
                pass

            raise

        finally:
            _safe_close(
                source_wb,
                save_changes=False,
            )
            _safe_close(
                output_wb,
                save_changes=False,
            )

            if excel is not None:
                try:
                    excel.ScreenUpdating = True
                    excel.EnableEvents = True
                    excel.Quit()
                except Exception:
                    pass

            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

