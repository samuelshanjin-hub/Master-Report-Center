from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any
import gc
import shutil
import tempfile
import time

from core.models.job_result import JobResult


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb", ".xls"}
HEADER_ROW = 1
XL_UP = -4162


@dataclass(frozen=True)
class HeaderAnalysis:
    destination_headers: tuple[str, ...]
    source_headers: tuple[str, ...]
    common_headers: tuple[str, ...]
    destination_sheet: str
    source_sheet: str
    source_sample_file: Path


def _normalise_header(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(
        str(value).replace("\r", " ").replace("\n", " ").split()
    ).casefold()


def _normalise_key_value(value: Any) -> tuple[str, Any]:
    if value is None:
        return ("blank", "")
    if isinstance(value, str):
        text = value.strip()
        if text == "":
            return ("blank", "")
        return ("text", text.casefold())
    if isinstance(value, bool):
        return ("bool", value)
    if isinstance(value, (int, float)):
        return ("number", round(float(value), 12))
    return (type(value).__name__, str(value).strip().casefold())


def _is_blank(value: Any) -> bool:
    return value is None or (
        isinstance(value, str) and value == ""
    )


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
            out = []
            for item in value:
                if isinstance(item, (tuple, list)):
                    out.append([item[0] if item else None])
                else:
                    out.append([item])
            return out
        return [[value]]

    if not isinstance(value, (tuple, list)):
        return [[value]]

    return [
        list(row) if isinstance(row, (tuple, list)) else [row]
        for row in value
    ]


def _excel_files(
    folder: Path,
    excluded: set[Path] | None = None,
) -> list[Path]:
    excluded = {
        p.resolve()
        for p in (excluded or set())
    }

    return sorted(
        p
        for p in folder.iterdir()
        if (
            p.is_file()
            and p.suffix.lower() in SUPPORTED_EXTENSIONS
            and not p.name.startswith("~$")
            and p.resolve() not in excluded
            and "_Updated_" not in p.stem
        )
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


def _safe_close(workbook: Any, save: bool = False, log=None) -> None:
    if workbook is None:
        return
    try:
        _call_with_com_retry(
            "closing workbook",
            lambda: workbook.Close(SaveChanges=save),
            log=log,
        )
    except Exception:
        pass


def _worksheet(workbook: Any, name: str):
    if name.strip():
        try:
            return workbook.Worksheets(name.strip())
        except Exception as exc:
            raise RuntimeError(
                f"Worksheet '{name.strip()}' was not found in "
                f"{Path(workbook.FullName).name}."
            ) from exc

    return workbook.Worksheets(1)


def _headers(ws: Any) -> tuple[list[str], dict[str, int]]:
    used = ws.UsedRange
    last_col = int(used.Column) + int(used.Columns.Count) - 1

    raw = ws.Range(
        ws.Cells(HEADER_ROW, 1),
        ws.Cells(HEADER_ROW, last_col),
    ).Value2

    matrix = _to_matrix(raw, 1, last_col)

    values = [
        "" if value is None else str(value)
        for value in matrix[0]
    ]

    normalised = [
        _normalise_header(value)
        for value in values
    ]

    while normalised and normalised[-1] == "":
        normalised.pop()
        values.pop()

    if not values or any(value == "" for value in normalised):
        raise RuntimeError(
            "Header row contains one or more blank headers."
        )

    if len(set(normalised)) != len(normalised):
        raise RuntimeError(
            "Header row contains duplicate headers."
        )

    return values, {
        normalised_header: index + 1
        for index, normalised_header
        in enumerate(normalised)
    }


def _last_row(ws: Any, header_map: dict[str, int]) -> int:
    maximum = HEADER_ROW

    for col in header_map.values():
        candidate = int(
            ws.Cells(
                ws.Rows.Count,
                col,
            ).End(XL_UP).Row
        )
        maximum = max(maximum, candidate)

    return maximum


def _elapsed_text(seconds: float) -> str:
    return str(timedelta(seconds=seconds))


class SynchronizeReportsService:
    """
    Project 003 Match & Update engine adapted to Master Report Center Core V2.

    Core behaviour intentionally follows Project 003 V0.9:
    - one Destination workbook
    - one folder of Source workbooks
    - one or more Match fields
    - one or more Write fields
    - first Destination row wins when duplicate destination keys exist
    - optional unmatched-row append
    - optional blank-source-value protection
    - original Destination is never overwritten
    """

    @staticmethod
    def _visibility_name(value: int) -> str:
        if int(value) == -1:
            return "Visible"
        if int(value) == 2:
            return "VeryHidden"
        return "Hidden"

    @staticmethod
    def _safe_bool_property(obj, name: str, default: bool = False) -> bool:
        try:
            return bool(getattr(obj, name))
        except Exception:
            return default

    @staticmethod
    def _is_sheet_protected(ws) -> bool:
        values = []
        for attr in (
            "ProtectContents",
            "ProtectDrawingObjects",
            "ProtectScenarios",
        ):
            try:
                values.append(bool(getattr(ws, attr)))
            except Exception:
                values.append(False)
        return any(values)

    def _capture_protection_settings(self, ws) -> dict:
        protected = self._is_sheet_protected(ws)
        state = {
            "is_protected": protected,
            "drawing_objects": self._safe_bool_property(
                ws, "ProtectDrawingObjects", False
            ),
            "contents": self._safe_bool_property(
                ws, "ProtectContents", False
            ),
            "scenarios": self._safe_bool_property(
                ws, "ProtectScenarios", False
            ),
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
                state[key] = self._safe_bool_property(
                    protection,
                    attr,
                    False,
                )

        return state

    def _unprotect_sheet(
        self,
        ws,
        password: str,
        protection_state: dict,
        log=None,
    ) -> None:
        if not protection_state.get("is_protected", False):
            return

        _call_with_com_retry(
            "unlocking Destination output worksheet",
            lambda: ws.Unprotect(Password=password or ""),
            log=log,
        )

        if self._is_sheet_protected(ws):
            raise RuntimeError(
                "The Destination output worksheet is still protected "
                "after the unlock attempt."
            )

    def _restore_sheet_protection(
        self,
        ws,
        protection_state: dict,
        password: str,
    ) -> None:
        if not protection_state.get("is_protected", False):
            return

        try:
            ws.Protect(
                Password=password or "",
                DrawingObjects=bool(
                    protection_state.get("drawing_objects", False)
                ),
                Contents=bool(
                    protection_state.get("contents", True)
                ),
                Scenarios=bool(
                    protection_state.get("scenarios", False)
                ),
                UserInterfaceOnly=False,
                AllowFormattingCells=bool(
                    protection_state.get("allow_formatting_cells", False)
                ),
                AllowFormattingColumns=bool(
                    protection_state.get("allow_formatting_columns", False)
                ),
                AllowFormattingRows=bool(
                    protection_state.get("allow_formatting_rows", False)
                ),
                AllowInsertingColumns=bool(
                    protection_state.get("allow_inserting_columns", False)
                ),
                AllowInsertingRows=bool(
                    protection_state.get("allow_inserting_rows", False)
                ),
                AllowInsertingHyperlinks=bool(
                    protection_state.get("allow_inserting_hyperlinks", False)
                ),
                AllowDeletingColumns=bool(
                    protection_state.get("allow_deleting_columns", False)
                ),
                AllowDeletingRows=bool(
                    protection_state.get("allow_deleting_rows", False)
                ),
                AllowSorting=bool(
                    protection_state.get("allow_sorting", False)
                ),
                AllowFiltering=bool(
                    protection_state.get("allow_filtering", False)
                ),
                AllowUsingPivotTables=bool(
                    protection_state.get("allow_using_pivot_tables", False)
                ),
            )
        except Exception as exc:
            raise RuntimeError(
                "Worksheet protection could not be restored on the "
                "synchronized output workbook."
            ) from exc

        if not self._is_sheet_protected(ws):
            raise RuntimeError(
                "Worksheet protection could not be verified after restoration."
            )

        if protection_state.get("enable_selection") is not None:
            try:
                ws.EnableSelection = int(
                    protection_state["enable_selection"]
                )
            except Exception:
                pass

    def _list_worksheets_from_workbook(self, workbook) -> list[dict]:
        items = []
        count = int(workbook.Worksheets.Count)

        for index in range(1, count + 1):
            ws = workbook.Worksheets(index)
            items.append(
                {
                    "name": str(ws.Name),
                    "visibility": self._visibility_name(
                        int(ws.Visible)
                    ),
                    "index": index,
                }
            )

        return items

    def get_destination_worksheets(
        self,
        destination_file: str,
    ) -> list[dict]:
        try:
            import pythoncom
            import win32com.client as win32
        except ImportError as exc:
            raise RuntimeError(
                "Synchronize Reports requires Microsoft Excel and pywin32 on Windows."
            ) from exc

        destination = Path(
            destination_file
        ).expanduser().resolve()

        if not destination.is_file():
            raise FileNotFoundError(
                f"Destination file not found: {destination}"
            )

        pythoncom.CoInitialize()
        excel = wb = None

        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False

            wb = _call_with_com_retry(
                "opening Destination workbook to list worksheets",
                lambda: excel.Workbooks.Open(
                    str(destination),
                    UpdateLinks=0,
                    ReadOnly=True,
                    AddToMru=False,
                    IgnoreReadOnlyRecommended=True,
                ),
            )

            return self._list_worksheets_from_workbook(wb)

        finally:
            _safe_close(wb)

            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass

            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def get_source_sample_worksheets(
        self,
        source_folder: str,
        destination_file: str = "",
    ) -> dict:
        try:
            import pythoncom
            import win32com.client as win32
        except ImportError as exc:
            raise RuntimeError(
                "Synchronize Reports requires Microsoft Excel and pywin32 on Windows."
            ) from exc

        folder = Path(
            source_folder
        ).expanduser().resolve()

        if not folder.is_dir():
            raise NotADirectoryError(
                f"Source folder not found: {folder}"
            )

        excluded = set()

        if destination_file:
            destination = Path(
                destination_file
            ).expanduser().resolve()
            if destination.exists():
                excluded.add(destination)

        files = _excel_files(
            folder,
            excluded=excluded,
        )

        if not files:
            raise RuntimeError(
                "No supported Excel source files were found."
            )

        pythoncom.CoInitialize()
        excel = wb = None
        last_error = None

        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False

            for path in files:
                try:
                    wb = _call_with_com_retry(
                        f"opening source workbook {path.name} to list worksheets",
                        lambda path=path: excel.Workbooks.Open(
                            str(path),
                            UpdateLinks=0,
                            ReadOnly=True,
                            AddToMru=False,
                            IgnoreReadOnlyRecommended=True,
                        ),
                    )

                    return {
                        "sample_file": path.name,
                        "worksheets": self._list_worksheets_from_workbook(wb),
                    }

                except Exception as exc:
                    last_error = exc

                finally:
                    _safe_close(wb)
                    wb = None

            raise RuntimeError(
                "No valid Source workbook could be opened."
                + (
                    f" Last error: {last_error}"
                    if last_error is not None
                    else ""
                )
            )

        finally:
            _safe_close(wb)

            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass

            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def get_destination_protection_info(
        self,
        destination_file: str,
        destination_sheet: str = "",
    ) -> dict:
        try:
            import pythoncom
            import win32com.client as win32
        except ImportError as exc:
            raise RuntimeError(
                "Synchronize Reports requires Microsoft Excel and pywin32 on Windows."
            ) from exc

        destination = Path(
            destination_file
        ).expanduser().resolve()

        pythoncom.CoInitialize()
        excel = wb = None

        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False

            wb = _call_with_com_retry(
                "opening Destination workbook for protection check",
                lambda: excel.Workbooks.Open(
                    str(destination),
                    UpdateLinks=0,
                    ReadOnly=True,
                    AddToMru=False,
                    IgnoreReadOnlyRecommended=True,
                ),
            )

            ws = _worksheet(
                wb,
                destination_sheet,
            )

            state = self._capture_protection_settings(ws)

            return {
                "is_protected": bool(state["is_protected"]),
                "sheet_name": str(ws.Name),
            }

        finally:
            _safe_close(wb)

            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass

            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def validate_destination_worksheet_password(
        self,
        destination_file: str,
        destination_sheet: str,
        password: str,
    ) -> bool:
        try:
            import pythoncom
            import win32com.client as win32
        except ImportError as exc:
            raise RuntimeError(
                "Synchronize Reports requires Microsoft Excel and pywin32 on Windows."
            ) from exc

        source_path = Path(
            destination_file
        ).expanduser().resolve()

        temp_dir = (
            Path(tempfile.gettempdir())
            / "GoldenFirReportCenter"
            / "SynchronizeReports"
            / "ProtectionCheck"
        )
        temp_dir.mkdir(
            parents=True,
            exist_ok=True,
        )

        temp_file = temp_dir / (
            f"PROTECTION_CHECK_"
            f"{datetime.now():%Y%m%d_%H%M%S_%f}_"
            f"{source_path.name}"
        )

        shutil.copy2(
            source_path,
            temp_file,
        )

        pythoncom.CoInitialize()
        excel = wb = None

        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False

            wb = _call_with_com_retry(
                "opening temporary Destination protection-check workbook",
                lambda: excel.Workbooks.Open(
                    str(temp_file),
                    UpdateLinks=0,
                    ReadOnly=False,
                    AddToMru=False,
                    IgnoreReadOnlyRecommended=True,
                ),
            )

            ws = _worksheet(
                wb,
                destination_sheet,
            )

            if not self._is_sheet_protected(ws):
                return True

            try:
                _call_with_com_retry(
                    "validating Destination worksheet password",
                    lambda: ws.Unprotect(
                        Password=password or ""
                    ),
                )
            except Exception:
                return False

            return not self._is_sheet_protected(
                ws
            )

        finally:
            _safe_close(wb)

            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass

            gc.collect()

            try:
                temp_file.unlink(
                    missing_ok=True
                )
            except Exception:
                pass

            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def analyze(
        self,
        destination_file: str,
        source_folder: str,
        destination_sheet: str = "",
        source_sheet: str = "",
        context=None,
    ) -> HeaderAnalysis:
        try:
            import pythoncom
            import win32com.client as win32
        except ImportError as exc:
            raise RuntimeError(
                "Synchronize Reports requires Microsoft Excel and pywin32 on Windows."
            ) from exc

        destination = Path(destination_file).expanduser().resolve()
        folder = Path(source_folder).expanduser().resolve()

        if not destination.is_file():
            raise FileNotFoundError(
                f"Destination file not found: {destination}"
            )

        if not folder.is_dir():
            raise NotADirectoryError(
                f"Source folder not found: {folder}"
            )

        files = _excel_files(
            folder,
            excluded={destination},
        )

        if not files:
            raise RuntimeError(
                "No supported Excel source files were found in the source folder."
            )

        def log(message: str) -> None:
            if context is not None:
                context.log(message)

        pythoncom.CoInitialize()

        excel = destination_wb = source_wb = None

        try:
            log("Starting Microsoft Excel for header analysis...")

            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False

            destination_wb = _call_with_com_retry(
                "opening Destination workbook for header analysis",
                lambda: excel.Workbooks.Open(
                    str(destination),
                    UpdateLinks=0,
                    ReadOnly=True,
                    AddToMru=False,
                ),
                log=log,
            )

            destination_ws = _worksheet(
                destination_wb,
                destination_sheet,
            )

            destination_headers, _ = _headers(
                destination_ws
            )

            source_wb = _call_with_com_retry(
                f"opening source workbook {files[0].name} for header analysis",
                lambda: excel.Workbooks.Open(
                    str(files[0]),
                    UpdateLinks=0,
                    ReadOnly=True,
                    AddToMru=False,
                    IgnoreReadOnlyRecommended=True,
                ),
                log=log,
            )

            source_ws = _worksheet(
                source_wb,
                source_sheet,
            )

            source_headers, _ = _headers(
                source_ws
            )

            source_header_map = {
                _normalise_header(header): header
                for header in source_headers
            }

            common_headers = tuple(
                header
                for header in destination_headers
                if _normalise_header(header)
                in source_header_map
            )

            log(
                f"Destination sheet: {destination_ws.Name}"
            )
            log(
                f"Source sample: {files[0].name}"
            )
            log(
                f"Source sheet: {source_ws.Name}"
            )
            log(
                f"Destination headers: {len(destination_headers)}"
            )
            log(
                f"Source headers: {len(source_headers)}"
            )
            log(
                f"Common headers available: {len(common_headers)}"
            )

            if context is not None:
                context.progress(
                    100,
                    "Header analysis completed.",
                )

            return HeaderAnalysis(
                destination_headers=tuple(
                    destination_headers
                ),
                source_headers=tuple(
                    source_headers
                ),
                common_headers=common_headers,
                destination_sheet=str(
                    destination_ws.Name
                ),
                source_sheet=str(
                    source_ws.Name
                ),
                source_sample_file=files[0],
            )

        finally:
            _safe_close(source_wb)
            _safe_close(destination_wb)

            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass

            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    @staticmethod
    def _write_log(
        log_file: Path,
        *,
        destination_file: Path,
        source_folder: Path,
        destination_sheet: str,
        source_sheet: str,
        match_headers: tuple[str, ...],
        write_headers: tuple[str, ...],
        append_unmatched: bool,
        ignore_blank_updates: bool,
        destination_worksheet_protected: bool,
        source_files_found: int,
        source_files_processed: int,
        files_skipped: int,
        skipped_details: list[tuple[str, str]],
        source_rows_read: int,
        matched_rows: int,
        updated_rows: int,
        matched_no_change: int,
        appended_rows: int,
        unmatched_rows: int,
        duplicate_destination_keys: int,
        output_file: Path | None,
        duration_seconds: float,
        status: str,
        error_message: str | None = None,
    ) -> None:
        lines = [
            "================================================",
            "Golden Fir Report Center - Synchronize Reports",
            "================================================",
            "",
            "Destination File:",
            str(destination_file),
            "",
            "Source Folder:",
            str(source_folder),
            "",
            "Destination Worksheet:",
            destination_sheet.strip() or "(First worksheet)",
            "",
            "Source Worksheet:",
            source_sheet.strip() or "(First worksheet)",
            "",
            "Match Fields:",
            " + ".join(match_headers),
            "",
            "Write Fields:",
            ", ".join(write_headers),
            "",
            "Append Unmatched Source Rows:",
            str(bool(append_unmatched)),
            "",
            "Ignore Blank Source Values:",
            str(bool(ignore_blank_updates)),
            "",
            "Destination Worksheet Protected:",
            str(bool(destination_worksheet_protected)),
            "",
            "Source Files Found:",
            str(source_files_found),
            "",
            "Source Files Processed:",
            str(source_files_processed),
            "",
            "Files Skipped:",
            str(files_skipped),
            "",
        ]

        if skipped_details:
            lines.append("Skipped File Details:")
            for filename, reason in skipped_details:
                lines.append(
                    f"- {filename} â€” {reason}"
                )
            lines.append("")

        lines.extend([
            "Source Rows Read:",
            str(source_rows_read),
            "",
            "Matched Rows:",
            str(matched_rows),
            "",
            "Updated Rows:",
            str(updated_rows),
            "",
            "Matched With No Write:",
            str(matched_no_change),
            "",
            "Appended Rows:",
            str(appended_rows),
            "",
            "Unmatched Rows:",
            str(unmatched_rows),
            "",
            "Duplicate Destination Keys:",
            str(duplicate_destination_keys),
            "",
            "Output File:",
            str(output_file) if output_file else "(None)",
            "",
            "Duration:",
            _elapsed_text(duration_seconds),
            "",
            "Status:",
            status,
        ])

        if error_message:
            lines.extend([
                "",
                "Error:",
                error_message,
            ])

        lines.extend([
            "",
            "================================================",
            "",
        ])

        log_file.write_text(
            "\n".join(lines),
            encoding="utf-8",
        )

    def run(
        self,
        destination_file: str,
        source_folder: str,
        destination_sheet: str,
        source_sheet: str,
        match_headers: tuple[str, ...],
        write_headers: tuple[str, ...],
        append_unmatched: bool,
        ignore_blank_updates: bool,
        worksheet_password: str = "",
        context=None,
    ) -> JobResult:
        try:
            import pythoncom
            import win32com.client as win32
        except ImportError as exc:
            raise RuntimeError(
                "Synchronize Reports requires Microsoft Excel and pywin32 on Windows."
            ) from exc

        started = time.perf_counter()

        destination = Path(
            destination_file
        ).expanduser().resolve()

        folder = Path(
            source_folder
        ).expanduser().resolve()

        if not destination.is_file():
            raise FileNotFoundError(
                f"Destination file not found: {destination}"
            )

        if not folder.is_dir():
            raise NotADirectoryError(
                f"Source folder not found: {folder}"
            )

        if not match_headers:
            raise RuntimeError(
                "Select at least one Match field."
            )

        if not write_headers:
            raise RuntimeError(
                "Select at least one Write field."
            )

        overlap = sorted(
            set(match_headers)
            & set(write_headers)
        )

        if overlap:
            raise RuntimeError(
                "The same field cannot be used as both "
                "a Match field and a Write field: "
                + ", ".join(overlap)
            )

        files = _excel_files(
            folder,
            excluded={destination},
        )

        if not files:
            raise RuntimeError(
                "No supported Excel source files were found."
            )

        stamp = datetime.now().strftime(
            "%Y-%m-%d_%H%M%S"
        )

        output_file = destination.with_name(
            f"{destination.stem}_Updated_"
            f"{stamp}{destination.suffix}"
        )

        log_file = destination.with_name(
            f"Synchronize_Log_{stamp}.txt"
        )

        shutil.copy2(
            destination,
            output_file,
        )

        def report(
            percent: int,
            message: str,
        ) -> None:
            safe = max(
                0,
                min(100, int(percent)),
            )
            context.progress(
                safe,
                message,
            )
            context.log(message)

        pythoncom.CoInitialize()

        excel = output_wb = source_wb = None

        processed = 0
        rows_read = 0
        matched = 0
        updated = 0
        appended = 0
        unmatched = 0
        skipped = 0
        duplicate_keys = 0
        matched_no_change = 0
        destination_protection_state = None

        skipped_details: list[
            tuple[str, str]
        ] = []

        try:
            report(
                1,
                "Starting Microsoft Excel...",
            )

            excel = win32.DispatchEx(
                "Excel.Application"
            )
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            excel.AskToUpdateLinks = False

            output_wb = _call_with_com_retry(
                f"opening output workbook {output_file.name}",
                lambda: excel.Workbooks.Open(
                    str(output_file),
                    UpdateLinks=0,
                    ReadOnly=False,
                    AddToMru=False,
                ),
                log=context.log,
            )

            output_ws = _worksheet(
                output_wb,
                destination_sheet,
            )

            destination_protection_state = (
                self._capture_protection_settings(
                    output_ws
                )
            )

            if destination_protection_state.get(
                "is_protected",
                False,
            ):
                report(
                    4,
                    "Destination worksheet is protected; "
                    "unlocking output copy...",
                )
                self._unprotect_sheet(
                    output_ws,
                    worksheet_password,
                    destination_protection_state,
                    log=context.log,
                )
                context.log(
                    "Destination worksheet unlocked in the copied "
                    "output workbook only."
                )

            destination_headers, destination_map = (
                _headers(output_ws)
            )

            match_norm = tuple(
                _normalise_header(header)
                for header in match_headers
            )

            write_norm = tuple(
                _normalise_header(header)
                for header in write_headers
            )

            for header in (
                match_norm + write_norm
            ):
                if header not in destination_map:
                    raise RuntimeError(
                        "Selected destination header "
                        f"not found: {header}"
                    )

            last_destination_row = _last_row(
                output_ws,
                destination_map,
            )

            index: dict[
                tuple[tuple[str, Any], ...],
                int,
            ] = {}

            duplicate_set = set()

            for row in range(
                HEADER_ROW + 1,
                last_destination_row + 1,
            ):
                key = tuple(
                    _normalise_key_value(
                        output_ws.Cells(
                            row,
                            destination_map[header],
                        ).Value2
                    )
                    for header in match_norm
                )

                if all(
                    part[0] == "blank"
                    for part in key
                ):
                    continue

                if key in index:
                    duplicate_set.add(key)
                else:
                    index[key] = row

            duplicate_keys = len(
                duplicate_set
            )

            report(
                8,
                (
                    f"Destination index created: "
                    f"{len(index):,} unique keys; "
                    f"{duplicate_keys:,} duplicate keys."
                ),
            )

            next_row = max(
                last_destination_row + 1,
                HEADER_ROW + 1,
            )

            for file_index, path in enumerate(
                files,
                start=1,
            ):
                base_percent = 10 + int(
                    (file_index - 1)
                    / len(files)
                    * 78
                )

                report(
                    base_percent,
                    f"Reading {path.name}...",
                )

                try:
                    source_wb = _call_with_com_retry(
                        f"opening source workbook {path.name}",
                        lambda path=path: excel.Workbooks.Open(
                            str(path),
                            UpdateLinks=0,
                            ReadOnly=True,
                            AddToMru=False,
                            IgnoreReadOnlyRecommended=True,
                        ),
                        log=context.log,
                    )

                    source_ws = _worksheet(
                        source_wb,
                        source_sheet,
                    )

                    (
                        source_headers,
                        source_map,
                    ) = _headers(source_ws)

                    for header in match_norm:
                        if header not in source_map:
                            original = match_headers[
                                match_norm.index(header)
                            ]
                            raise RuntimeError(
                                f"Match header missing: {original}"
                            )

                    available_write = [
                        header
                        for header in write_norm
                        if (
                            header in source_map
                            and header
                            in destination_map
                        )
                    ]

                    source_last_row = _last_row(
                        source_ws,
                        source_map,
                    )

                    for source_row in range(
                        HEADER_ROW + 1,
                        source_last_row + 1,
                    ):
                        key = tuple(
                            _normalise_key_value(
                                source_ws.Cells(
                                    source_row,
                                    source_map[header],
                                ).Value2
                            )
                            for header in match_norm
                        )

                        if all(
                            part[0] == "blank"
                            for part in key
                        ):
                            continue

                        rows_read += 1
                        destination_row = index.get(
                            key
                        )

                        if destination_row is not None:
                            matched += 1
                            changed = False

                            for header in available_write:
                                value = source_ws.Cells(
                                    source_row,
                                    source_map[header],
                                ).Value2

                                if (
                                    ignore_blank_updates
                                    and _is_blank(value)
                                ):
                                    continue

                                # Project 003 established behaviour:
                                # values-only writeback so Destination
                                # formatting remains intact.
                                _call_with_com_retry(
                                    f"writing matched field {header}",
                                    lambda destination_row=destination_row, header=header, value=value: setattr(
                                        output_ws.Cells(
                                            destination_row,
                                            destination_map[header],
                                        ),
                                        "Value2",
                                        value,
                                    ),
                                    log=context.log,
                                )

                                changed = True

                            if changed:
                                updated += 1
                            else:
                                matched_no_change += 1

                        elif append_unmatched:
                            # Append every Source field that has a
                            # same-name Destination header.
                            for (
                                normalised_header,
                                source_col,
                            ) in source_map.items():
                                if (
                                    normalised_header
                                    in destination_map
                                ):
                                    output_ws.Cells(
                                        next_row,
                                        destination_map[
                                            normalised_header
                                        ],
                                    ).Value2 = (
                                        source_ws.Cells(
                                            source_row,
                                            source_col,
                                        ).Value2
                                    )

                            index[key] = next_row
                            next_row += 1
                            appended += 1

                        else:
                            unmatched += 1

                    processed += 1

                except Exception as exc:
                    skipped += 1
                    reason = str(exc)

                    skipped_details.append(
                        (
                            path.name,
                            reason,
                        )
                    )

                    report(
                        base_percent,
                        f"Skipped {path.name}: {reason}",
                    )

                finally:
                    _safe_close(source_wb)
                    source_wb = None

            report(
                88,
                (
                    f"Matched rows: {matched:,}; "
                    f"updated rows: {updated:,}; "
                    f"matched with no write: "
                    f"{matched_no_change:,}; "
                    f"appended rows: {appended:,}."
                ),
            )

            if (
                destination_protection_state is not None
                and destination_protection_state.get(
                    "is_protected",
                    False,
                )
            ):
                report(
                    91,
                    "Restoring worksheet protection on updated Destination copy...",
                )
                self._restore_sheet_protection(
                    output_ws,
                    destination_protection_state,
                    worksheet_password,
                )
                context.log(
                    "Worksheet protection restored and verified "
                    "on updated Destination copy."
                )

            report(
                92,
                "Saving updated Destination copy...",
            )

            _call_with_com_retry(
                f"saving output workbook {output_file.name}",
                lambda: output_wb.Save(),
                log=context.log,
            )
            _safe_close(
                output_wb,
                save=True,
                log=context.log,
            )
            output_wb = None

            elapsed = (
                time.perf_counter()
                - started
            )

            self._write_log(
                log_file,
                destination_file=destination,
                source_folder=folder,
                destination_sheet=destination_sheet,
                source_sheet=source_sheet,
                match_headers=match_headers,
                write_headers=write_headers,
                append_unmatched=append_unmatched,
                ignore_blank_updates=ignore_blank_updates,
                destination_worksheet_protected=bool(
                    destination_protection_state
                    and destination_protection_state.get(
                        "is_protected",
                        False,
                    )
                ),
                source_files_found=len(files),
                source_files_processed=processed,
                files_skipped=skipped,
                skipped_details=skipped_details,
                source_rows_read=rows_read,
                matched_rows=matched,
                updated_rows=updated,
                matched_no_change=matched_no_change,
                appended_rows=appended,
                unmatched_rows=unmatched,
                duplicate_destination_keys=duplicate_keys,
                output_file=output_file,
                duration_seconds=elapsed,
                status="Completed Successfully",
            )

            context.log("")
            context.log(
                "========== Synchronization Summary =========="
            )
            context.log(
                f"Source files found: {len(files)}"
            )
            context.log(
                f"Source files processed: {processed}"
            )
            context.log(
                f"Files skipped: {skipped}"
            )
            context.log(
                f"Source rows read: {rows_read}"
            )
            context.log(
                f"Matched rows: {matched}"
            )
            context.log(
                f"Updated rows: {updated}"
            )
            context.log(
                f"Matched with no write: "
                f"{matched_no_change}"
            )
            context.log(
                f"Appended rows: {appended}"
            )
            context.log(
                f"Unmatched rows: {unmatched}"
            )
            context.log(
                "Duplicate destination keys: "
                f"{duplicate_keys}"
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
                "Synchronization completed.",
            )

            return JobResult(
                success=True,
                message=(
                    f"Synchronization completed; "
                    f"{updated} matched row(s) updated, "
                    f"{appended} unmatched row(s) appended."
                ),
                output_files=[
                    str(output_file),
                    str(log_file),
                ],
                processed_count=processed,
                skipped_count=skipped,
                warning_count=(
                    skipped
                    + duplicate_keys
                ),
            )

        except Exception as exc:
            elapsed = (
                time.perf_counter()
                - started
            )

            try:
                self._write_log(
                    log_file,
                    destination_file=destination,
                    source_folder=folder,
                    destination_sheet=destination_sheet,
                    source_sheet=source_sheet,
                    match_headers=match_headers,
                    write_headers=write_headers,
                    append_unmatched=append_unmatched,
                    ignore_blank_updates=ignore_blank_updates,
                    destination_worksheet_protected=bool(
                        destination_protection_state
                        and destination_protection_state.get(
                            "is_protected",
                            False,
                        )
                    ),
                    source_files_found=len(files),
                    source_files_processed=processed,
                    files_skipped=skipped,
                    skipped_details=skipped_details,
                    source_rows_read=rows_read,
                    matched_rows=matched,
                    updated_rows=updated,
                    matched_no_change=matched_no_change,
                    appended_rows=appended,
                    unmatched_rows=unmatched,
                    duplicate_destination_keys=duplicate_keys,
                    output_file=output_file,
                    duration_seconds=elapsed,
                    status="Failed",
                    error_message=str(exc),
                )

                context.log(
                    f"Persistent failure log: {log_file}"
                )

            except Exception:
                pass

            raise

        finally:
            _safe_close(source_wb)
            _safe_close(output_wb)

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

