from __future__ import annotations

import gc
import os
import re
import shutil
import tempfile
import time
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

import pandas as pd

from core.models.job_result import JobResult


XL_CELL_TYPE_VISIBLE = 12
XL_UP = -4162
XL_TO_LEFT = -4159
XL_CELL_TYPE_ALL_VALIDATION = -4174
XL_VALIDATE_LIST = 3


def _safe_filename(value) -> str:
    if value is None:
        return "Unknown"
    text = str(value).strip()
    if not text:
        return "Unknown"
    text = re.sub(r'[\\/:*?"<>|]', "_", text)
    return text[:180]


def _normalise(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).replace("\u00a0", " ").strip().casefold()


def _require_pywin32():
    try:
        import pythoncom  # type: ignore
        import win32com.client as win32  # type: ignore
        from pywintypes import com_error  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "Split Reports requires Microsoft Excel and pywin32 on Windows. "
            "Install with: python -m pip install -r requirements-windows-com.txt"
        ) from exc
    return pythoncom, win32, com_error


class SplitReportsService:
    """Publish Split Reports service backed by the proven Excel reverse-filter approach."""

    _COM_BUSY_HRESULTS = {
        -2147418111,  # RPC_E_CALL_REJECTED: Call was rejected by callee.
        -2147417846,  # RPC_E_SERVERCALL_RETRYLATER: Server call retry later.
    }

    @classmethod
    def _is_retryable_com_busy_error(cls, exc: Exception) -> bool:
        """Return True only for transient Excel/COM busy-call failures."""
        hresult = getattr(exc, "hresult", None)

        if hresult in cls._COM_BUSY_HRESULTS:
            return True

        args = getattr(exc, "args", ()) or ()
        if args:
            first = args[0]
            if isinstance(first, int) and first in cls._COM_BUSY_HRESULTS:
                return True

        message = str(exc).casefold()
        busy_markers = (
            "call was rejected by callee",
            "server call retry later",
            "application is busy",
            "message filter indicated that the application is busy",
        )
        return any(marker in message for marker in busy_markers)

    def _call_with_com_retry(
        self,
        operation: str,
        func,
        log: Optional[Callable[[str], None]] = None,
        attempts: int = 5,
        initial_delay: float = 0.5,
    ):
        """
        Run one Excel COM operation with bounded retry for transient busy errors.

        Non-busy COM errors are raised immediately. Busy errors are retried with
        a short increasing delay so a temporary Excel save/copy/delete state does
        not abort the entire split job.
        """
        attempts = max(int(attempts), 1)

        for attempt in range(1, attempts + 1):
            try:
                return func()
            except Exception as exc:
                if not self._is_retryable_com_busy_error(exc) or attempt >= attempts:
                    raise

                delay = initial_delay * attempt

                if log is not None:
                    log(
                        f"Excel is busy during {operation}. "
                        f"Retrying COM operation ({attempt + 1}/{attempts}) "
                        f"in {delay:.1f}s..."
                    )

                # Give the Excel STA a brief chance to finish the previous call.
                try:
                    import pythoncom  # type: ignore
                    pythoncom.PumpWaitingMessages()
                except Exception:
                    pass

                time.sleep(delay)

    @staticmethod
    def build_output_filename(
        split_value,
        custom_suffix: str = "",
        add_date_suffix: bool = False,
        extension: str = ".xlsx",
    ) -> str:
        """
        Build:
            <Split Value> - <Custom Suffix> - <DD-MM-YYYY><source extension>

        The output extension must match the workbook format being copied.
        This is critical for .xlsm / .xls sources because the Split engine uses
        a native workbook copy rather than converting formats.
        """
        parts = [_safe_filename(split_value)]

        suffix = str(custom_suffix or "").strip()
        if suffix:
            safe_suffix = _safe_filename(suffix)
            if safe_suffix:
                parts.append(safe_suffix)

        if add_date_suffix:
            parts.append(datetime.now().strftime("%d-%m-%Y"))

        ext = str(extension or ".xlsx").strip().lower()
        if not ext.startswith("."):
            ext = "." + ext

        # Formats currently accepted by the Split UI / engine.
        if ext not in {".xlsx", ".xlsm", ".xls"}:
            ext = ".xlsx"

        return " - ".join(parts) + ext

    @staticmethod
    def _is_sheet_protected(ws) -> bool:
        """Return True when any worksheet protection area is active."""
        values = []
        for attr in ("ProtectContents", "ProtectDrawingObjects", "ProtectScenarios"):
            try:
                values.append(bool(getattr(ws, attr)))
            except Exception:
                values.append(False)
        return any(values)

    @staticmethod
    def _safe_bool_property(obj, name: str, default: bool = False) -> bool:
        try:
            return bool(getattr(obj, name))
        except Exception:
            return default

    def _capture_protection_settings(self, ws) -> dict:
        """
        Capture the worksheet protection state and the Excel protection options
        that can be read back through COM.

        Cell/row/column Locked properties are not rebuilt. They are already
        preserved by copying the workbook and remain attached to surviving cells.
        """
        protected = self._is_sheet_protected(ws)

        state = {
            "is_protected": protected,
            "drawing_objects": self._safe_bool_property(
                ws, "ProtectDrawingObjects", False
            ),
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
            state["enable_selection"] = None

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

    def _unprotect_sheet(
        self,
        ws,
        password: str,
        protection_state: dict,
    ) -> None:
        if not protection_state.get("is_protected", False):
            return

        try:
            ws.Unprotect(Password=password or "")
        except Exception as exc:
            raise RuntimeError(
                "The worksheet protection password was not accepted."
            ) from exc

        if self._is_sheet_protected(ws):
            raise RuntimeError(
                "The worksheet is still protected after the unlock attempt."
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
                DrawingObjects=bool(protection_state.get("drawing_objects", False)),
                Contents=bool(protection_state.get("contents", True)),
                Scenarios=bool(protection_state.get("scenarios", False)),
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
                AllowSorting=bool(protection_state.get("allow_sorting", False)),
                AllowFiltering=bool(
                    protection_state.get("allow_filtering", False)
                ),
                AllowUsingPivotTables=bool(
                    protection_state.get("allow_using_pivot_tables", False)
                ),
            )
        except Exception as exc:
            raise RuntimeError(
                "Worksheet protection could not be restored. "
                "The output workbook was not accepted as complete."
            ) from exc

        if not self._is_sheet_protected(ws):
            raise RuntimeError(
                "Worksheet protection could not be verified after restoration."
            )

        enable_selection = protection_state.get("enable_selection")
        if enable_selection is not None:
            try:
                ws.EnableSelection = int(enable_selection)
            except Exception:
                # EnableSelection is secondary to the actual protection state.
                pass

    def _verify_saved_worksheet_password(
        self,
        excel,
        output_file: Path,
        sheet_name: str,
        protection_state: dict,
        password: str,
        log: Optional[Callable[[str], None]] = None,
    ) -> None:
        """
        Verify the protection that was actually persisted to disk.

        For a password-protected source worksheet, the saved split output must:
        1) still be protected after reopening;
        2) reject a blank password; and
        3) accept the original session password.

        Verification is performed against a temporary copy of the completed
        output workbook, so the accepted output file itself is never modified
        by the password test.
        """
        if not protection_state.get("is_protected", False):
            return

        output_file = Path(output_file)
        verify_dir = (
            Path(tempfile.gettempdir())
            / "GoldenFirReportCenter"
            / "SplitReports"
            / "PasswordVerify"
        )
        verify_dir.mkdir(parents=True, exist_ok=True)
        verify_file = verify_dir / (
            f"VERIFY_{datetime.now():%Y%m%d_%H%M%S_%f}_{output_file.name}"
        )

        verify_wb = None
        verify_ws = None

        try:
            shutil.copy2(output_file, verify_file)

            verify_wb = self._call_with_com_retry(
                f"opening password-verification copy for {output_file.name}",
                lambda: excel.Workbooks.Open(
                    str(verify_file),
                    UpdateLinks=0,
                    ReadOnly=False,
                ),
                log=log,
            )
            verify_ws = verify_wb.Worksheets(sheet_name)

            if not self._is_sheet_protected(verify_ws):
                raise RuntimeError(
                    f"{output_file.name}: worksheet protection was not preserved "
                    "after saving. The output workbook was not accepted."
                )

            expected_password = str(password or "")

            # If the original protection used a non-blank password, a blank
            # Unprotect attempt must NOT be able to remove the saved protection.
            if expected_password:
                blank_unlocked = False
                try:
                    self._call_with_com_retry(
                        f"testing blank worksheet password for {output_file.name}",
                        lambda: verify_ws.Unprotect(Password=""),
                        log=log,
                    )
                    blank_unlocked = not self._is_sheet_protected(verify_ws)
                except Exception:
                    # Expected result for a password-protected worksheet.
                    blank_unlocked = False

                if blank_unlocked:
                    raise RuntimeError(
                        f"{output_file.name}: worksheet protection was restored "
                        "without the original password. Split stopped because the "
                        "saved output accepts a blank password."
                    )

                # The original session password must successfully unlock the
                # persisted protection.
                try:
                    self._call_with_com_retry(
                        f"verifying worksheet password for {output_file.name}",
                        lambda: verify_ws.Unprotect(
                            Password=expected_password
                        ),
                        log=log,
                    )
                except Exception as exc:
                    raise RuntimeError(
                        f"{output_file.name}: the saved worksheet protection does "
                        "not accept the original worksheet password. The output "
                        "workbook was not accepted."
                    ) from exc

                if self._is_sheet_protected(verify_ws):
                    raise RuntimeError(
                        f"{output_file.name}: the saved worksheet remained "
                        "protected after verification with the original password."
                    )

                if log is not None:
                    log(
                        f"{output_file.name}: worksheet password protection verified."
                    )
            elif log is not None:
                log(
                    f"{output_file.name}: worksheet protection verified "
                    "(source protection uses a blank password)."
                )

        finally:
            try:
                if verify_wb is not None:
                    self._call_with_com_retry(
                        "closing password-verification workbook",
                        lambda: verify_wb.Close(SaveChanges=False),
                        log=log,
                    )
            except Exception:
                pass

            verify_ws = None
            verify_wb = None

            try:
                verify_file.unlink(missing_ok=True)
            except Exception:
                pass

    def get_worksheet_protection_info(
        self,
        source_file: str,
        sheet_name: str,
        use_open_excel_view: bool = False,
    ) -> dict:
        """
        Read worksheet protection without changing the source workbook.
        """
        pythoncom, win32, _ = _require_pywin32()
        pythoncom.CoInitialize()

        excel = None
        wb = None
        ws = None
        attached = False

        try:
            if use_open_excel_view:
                try:
                    excel = win32.GetObject(None, "Excel.Application")
                except Exception as exc:
                    raise RuntimeError(
                        "Microsoft Excel is not open. Open the selected workbook first."
                    ) from exc

                wb = self._find_open_workbook(excel, source_file)
                if wb is None:
                    raise RuntimeError(
                        "Please open the selected workbook in Excel first."
                    )
                attached = True
            else:
                excel = win32.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                excel.ScreenUpdating = False
                excel.EnableEvents = False
                wb = self._call_with_com_retry(
                    "opening workbook for protection check",
                    lambda: excel.Workbooks.Open(
                        str(Path(source_file).resolve()),
                        UpdateLinks=0,
                        ReadOnly=True,
                    ),
                )

            ws = wb.Worksheets(sheet_name)
            state = self._capture_protection_settings(ws)

            return {
                "is_protected": bool(state["is_protected"]),
                "protect_contents": bool(state["contents"]),
                "protect_drawing_objects": bool(state["drawing_objects"]),
                "protect_scenarios": bool(state["scenarios"]),
            }

        finally:
            if not attached:
                try:
                    if wb is not None:
                        self._call_with_com_retry(
                            "closing workbook after protection check",
                            lambda: wb.Close(SaveChanges=False),
                        )
                except Exception:
                    pass
                try:
                    if excel is not None:
                        excel.Quit()
                except Exception:
                    pass

            ws = None
            wb = None
            excel = None
            gc.collect()

            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def validate_worksheet_password(
        self,
        source_file: str,
        sheet_name: str,
        password: str,
        use_open_excel_view: bool = False,
    ) -> bool:
        """
        Validate a worksheet password against a temporary workbook copy.

        The source workbook is never unprotected or saved. This method is also
        used to test a blank password before the UI decides whether to prompt.
        """
        pythoncom, win32, _ = _require_pywin32()
        pythoncom.CoInitialize()

        temp_dir = (
            Path(tempfile.gettempdir())
            / "GoldenFirReportCenter"
            / "SplitReports"
            / "ProtectionCheck"
        )
        temp_dir.mkdir(parents=True, exist_ok=True)
        temp_file = temp_dir / (
            f"PROTECTION_CHECK_{datetime.now():%Y%m%d_%H%M%S_%f}_"
            f"{Path(source_file).name}"
        )

        attached_excel = None
        source_wb = None
        test_excel = None
        test_wb = None
        test_ws = None

        try:
            if use_open_excel_view:
                try:
                    attached_excel = win32.GetObject(None, "Excel.Application")
                except Exception as exc:
                    raise RuntimeError(
                        "Microsoft Excel is not open. Open the selected workbook first."
                    ) from exc

                source_wb = self._find_open_workbook(
                    attached_excel,
                    source_file,
                )
                if source_wb is None:
                    raise RuntimeError(
                        "Please open the selected workbook in Excel first."
                    )
                self._call_with_com_retry(
                    "creating temporary protection-check copy",
                    lambda: source_wb.SaveCopyAs(str(temp_file)),
                )
            else:
                shutil.copy2(source_file, temp_file)

            test_excel = win32.DispatchEx("Excel.Application")
            test_excel.Visible = False
            test_excel.DisplayAlerts = False
            test_excel.ScreenUpdating = False
            test_excel.EnableEvents = False

            test_wb = self._call_with_com_retry(
                "opening temporary protection-check workbook",
                lambda: test_excel.Workbooks.Open(
                    str(temp_file),
                    UpdateLinks=0,
                    ReadOnly=False,
                ),
            )
            test_ws = test_wb.Worksheets(sheet_name)

            state = self._capture_protection_settings(test_ws)
            if not state["is_protected"]:
                return True

            try:
                self._call_with_com_retry(
                    "unlocking temporary worksheet",
                    lambda: test_ws.Unprotect(Password=password or ""),
                )
            except Exception:
                return False

            return not self._is_sheet_protected(test_ws)

        finally:
            try:
                if test_wb is not None:
                    self._call_with_com_retry(
                        "closing temporary protection-check workbook",
                        lambda: test_wb.Close(SaveChanges=False),
                    )
            except Exception:
                pass

            try:
                if test_excel is not None:
                    test_excel.Quit()
            except Exception:
                pass

            test_ws = None
            test_wb = None
            test_excel = None
            source_wb = None
            attached_excel = None
            gc.collect()

            try:
                temp_file.unlink(missing_ok=True)
            except Exception:
                pass

            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def get_sheet_names(self, source_file: str) -> list[str]:
        path = Path(source_file)
        if not path.is_file():
            raise FileNotFoundError(f"File not found: {source_file}")
        return list(pd.ExcelFile(path).sheet_names)

    def detect_header_row(
        self,
        source_file: str,
        sheet_name: str,
        scan_rows: int = 20,
    ) -> int:
        """
        Detect the most likely field-header row within the first scan_rows rows.

        The detector deliberately favours rows containing many distinct text
        labels and penalises title/section rows containing only one or two cells.
        The UI always provides a manual override because Excel templates vary.
        """
        preview = pd.read_excel(
            source_file,
            sheet_name=sheet_name,
            header=None,
            nrows=max(int(scan_rows), 1),
            dtype=object,
        )

        if preview.empty:
            return 1

        best_row = 1
        best_score = float("-inf")

        for zero_index, row in preview.iterrows():
            values = []
            for value in row.tolist():
                if value is None:
                    continue
                try:
                    if pd.isna(value):
                        continue
                except Exception:
                    pass
                text = str(value).replace("\u00a0", " ").strip()
                if text:
                    values.append(value)

            nonblank = len(values)
            if nonblank == 0:
                continue

            text_count = 0
            numeric_count = 0
            distinct_text = set()

            for value in values:
                if isinstance(value, str):
                    text = value.replace("\u00a0", " ").strip()
                    if text:
                        text_count += 1
                        distinct_text.add(text.casefold())
                elif isinstance(value, (int, float)) and not isinstance(value, bool):
                    numeric_count += 1
                else:
                    text = str(value).strip()
                    if text:
                        text_count += 1
                        distinct_text.add(text.casefold())

            # Header rows are normally label-heavy and diverse.
            score = (
                nonblank * 3.0
                + text_count * 2.5
                + len(distinct_text) * 1.5
                - numeric_count * 1.5
            )

            # Strongly de-prioritise report titles / section-label rows.
            if nonblank <= 2:
                score -= 18.0
            elif nonblank <= 4:
                score -= 6.0

            # Earlier rows win close scores.
            score -= zero_index * 0.05

            if score > best_score:
                best_score = score
                best_row = int(zero_index) + 1

        return max(best_row, 1)

    def get_columns(
        self,
        source_file: str,
        sheet_name: str,
        header_row: int = 1,
    ) -> list[str]:
        header_row = max(int(header_row), 1)
        df = pd.read_excel(
            source_file,
            sheet_name=sheet_name,
            header=header_row - 1,
            nrows=0,
        )
        return [str(c) for c in df.columns]

    def analyze(
        self,
        source_file: str,
        sheet_name: str,
        split_column: str,
        use_open_excel_view: bool = False,
        worksheet_password: str = "",
        header_row: int = 1,
    ) -> dict:
        """
        Full Workbook analysis reads the file directly.

        Current View analysis uses the same Excel-native temporary view conversion
        that Start Splitting uses. No Python-side visible-row inference is used.
        """
        if use_open_excel_view:
            temp_file = None
            try:
                temp_file, _ = self._make_visible_view_master(
                    original_source=source_file,
                    sheet_name=sheet_name,
                    log=lambda _message: None,
                    worksheet_password=worksheet_password,
                    expected_analysis_rows=None,
                    header_row=header_row,
                )
                result = self._analyze_full_workbook(
                    str(temp_file),
                    sheet_name,
                    split_column,
                    include_protection_check=False,
                    header_row=header_row,
                )
                result["source_mode"] = "Current Excel Workbook View"
                result["worksheet_protected"] = bool(
                    self.get_worksheet_protection_info(
                        source_file=source_file,
                        sheet_name=sheet_name,
                        use_open_excel_view=True,
                    )["is_protected"]
                )
                return result
            finally:
                if temp_file is not None:
                    try:
                        Path(temp_file).unlink(missing_ok=True)
                    except Exception:
                        pass

        return self._analyze_full_workbook(
            source_file,
            sheet_name,
            split_column,
            include_protection_check=True,
            header_row=header_row,
        )

    def _analyze_full_workbook(
        self,
        source_file: str,
        sheet_name: str,
        split_column: str,
        include_protection_check: bool = True,
        header_row: int = 1,
    ) -> dict:
        header_row = max(int(header_row), 1)
        df = pd.read_excel(
            source_file,
            sheet_name=sheet_name,
            header=header_row - 1,
            dtype=object,
        )

        if split_column not in df.columns:
            raise ValueError(f"Split column not found: {split_column}")

        cleaned = df[split_column].fillna("").astype(str).str.strip()
        blank_rows = int((cleaned == "").sum())
        nonblank = cleaned[cleaned != ""]
        unique_values = sorted(nonblank.unique().tolist(), key=_normalise)

        worksheet_protected = False
        if include_protection_check:
            try:
                protection_info = self.get_worksheet_protection_info(
                    source_file=source_file,
                    sheet_name=sheet_name,
                    use_open_excel_view=False,
                )
                worksheet_protected = bool(protection_info["is_protected"])
            except Exception:
                worksheet_protected = False

        return {
            "worksheet": sheet_name,
            "rows": int(len(df)),
            "columns": int(len(df.columns)),
            "output_files": int(len(unique_values)),
            "blank_rows": blank_rows,
            "source_mode": "Full Workbook Data",
            "example_value": str(unique_values[0]).strip() if unique_values else "",
            "worksheet_protected": worksheet_protected,
            "header_row": header_row,
        }


    @staticmethod
    def _create_output_folder(base_output: str) -> Path:
        base = Path(base_output)
        if not base.exists():
            base.mkdir(parents=True, exist_ok=True)
        date = datetime.now().strftime("%d%m%Y")
        stem = f"Report Split_{date}"
        out = base / stem
        i = 1
        while out.exists():
            out = base / f"{stem}_{i:02d}"
            i += 1
        out.mkdir(parents=True, exist_ok=False)
        return out

    @staticmethod
    def _is_network_or_cloud(path: Path) -> bool:
        text = str(path).lower()
        if str(path).startswith("\\\\"):
            return True
        if path.drive and path.drive.upper() not in {"C:", "D:"}:
            return True
        return any(k in text for k in ("onedrive", "sharepoint", "dropbox", "google drive"))

    def _prepare_local_source(self, source_file: str, log: Callable[[str], None]) -> tuple[Path, bool]:
        source = Path(source_file).resolve()
        if not source.exists():
            raise FileNotFoundError(f"File not found: {source}")
        if not self._is_network_or_cloud(source):
            log("Source type: LOCAL")
            return source, False
        temp_dir = Path(tempfile.gettempdir()) / "GoldenFirReportCenter" / "SplitReports"
        temp_dir.mkdir(parents=True, exist_ok=True)
        target = temp_dir / f"{source.stem}_{datetime.now():%Y%m%d_%H%M%S}{source.suffix}"
        log("Source type: NETWORK / SYNCED CLOUD")
        log(f"Copying source to local temporary file: {target}")
        shutil.copy2(source, target)
        return target, True

    @staticmethod
    def _find_open_workbook(excel, source_file: str):
        target = os.path.abspath(source_file).lower()
        target_name = os.path.basename(target)
        for wb in excel.Workbooks:
            try:
                full = os.path.abspath(wb.FullName).lower()
                if full == target or os.path.basename(full) == target_name:
                    return wb
            except Exception:
                continue
        return None

    @staticmethod
    def _current_view_scope(ws, header_row: int = 1):
        """
        Return the Excel range whose current filter/hide state defines the view.

        header_row is explicit so report titles/notes above the field headers are
        preserved and are never treated as data rows.
        """
        header_row = max(int(header_row), 1)

        try:
            if bool(ws.AutoFilterMode):
                rng = ws.AutoFilter.Range
                filter_first_col = int(rng.Column)
                filter_last_col = filter_first_col + int(rng.Columns.Count) - 1
                filter_last_row = int(rng.Row) + int(rng.Rows.Count) - 1
                if int(rng.Row) == header_row:
                    return (
                        header_row,
                        filter_last_row,
                        filter_first_col,
                        filter_last_col,
                        f"AutoFilter.Range {rng.Address}",
                    )
        except Exception:
            pass

        used = ws.UsedRange
        first_col = int(used.Column)
        last_col = first_col + int(used.Columns.Count) - 1
        last_row = int(used.Row) + int(used.Rows.Count) - 1
        return (
            header_row,
            max(last_row, header_row),
            first_col,
            last_col,
            f"UsedRange {used.Address} with Header Row {header_row}",
        )

    @staticmethod
    def _first_nonhidden_column(ws, first_col: int, last_col: int) -> int:
        for col in range(first_col, last_col + 1):
            try:
                if not bool(ws.Columns(col).Hidden):
                    return col
            except Exception:
                return col
        return first_col

    def _convert_current_view_in_place(
        self,
        ws,
        log: Callable[[str], None],
        header_row: int = 1,
    ) -> tuple[int, int]:
        """
        Excel-native Current View conversion.

        Excel itself decides which rows are visible while the original
        Filter/Hide state is still active. Those rows are marked KEEP in a
        helper column.

        FIX11 change:
        Instead of asking Excel to delete all filtered DELETE rows through one
        SpecialCells(...).EntireRow.Delete() operation, read the helper markers
        back in one COM call, build contiguous DELETE row blocks, clear the
        original filter, then delete those blocks bottom-up.

        This preserves the VBA-equivalent visibility decision while avoiding
        Excel's fragile bulk deletion of filtered rows.
        """
        header_row, last_row, first_col, last_col, scope = self._current_view_scope(
            ws,
            header_row=header_row,
        )

        if last_row <= header_row:
            return 0, 0

        reference_col = self._first_nonhidden_column(ws, first_col, last_col)
        helper_col = self._get_safe_helper_column(
            ws,
            minimum_column=last_col,
            log=log,
        )

        log(f"Current view scope: {scope}")
        log(
            f"Current view data rows in scope: "
            f"{max(last_row - header_row, 0):,}"
        )

        ws.Cells(header_row, helper_col).Value = "__GF_VIEW_KEEP__"

        total_rows = last_row - header_row

        # Default every data row to DELETE in a single COM write.
        ws.Range(
            ws.Cells(header_row + 1, helper_col),
            ws.Cells(last_row, helper_col),
        ).Value = [("DELETE",)] * total_rows

        # While the original Excel view is still active, ask Excel itself for
        # the visible areas in one non-hidden reference column.
        reference_range = ws.Range(
            ws.Cells(header_row + 1, reference_col),
            ws.Cells(last_row, reference_col),
        )

        try:
            visible = reference_range.SpecialCells(XL_CELL_TYPE_VISIBLE)
        except Exception as exc:
            raise RuntimeError(
                "Excel could not identify the currently visible rows in the "
                "temporary workbook view."
            ) from exc

        visible_count = 0

        for area in visible.Areas:
            area_start = max(int(area.Row), header_row + 1)
            area_end = min(
                int(area.Row) + int(area.Rows.Count) - 1,
                last_row,
            )

            if area_end < area_start:
                continue

            count = area_end - area_start + 1
            visible_count += count

            ws.Range(
                ws.Cells(area_start, helper_col),
                ws.Cells(area_end, helper_col),
            ).Value = [("KEEP",)] * count

        log(f"Excel-native visible rows marked: {visible_count:,}")

        if visible_count <= 0:
            raise RuntimeError(
                "Excel reported zero visible data rows in the current view."
            )

        # Read all KEEP/DELETE markers before changing the view.
        markers = ws.Range(
            ws.Cells(header_row + 1, helper_col),
            ws.Cells(last_row, helper_col),
        ).Value

        if markers is None:
            markers = ()
        elif not isinstance(markers, tuple):
            markers = ((markers,),)

        delete_ranges: list[tuple[int, int]] = []
        range_start = None

        for offset, item in enumerate(markers):
            row_number = header_row + 1 + offset
            value = item[0] if isinstance(item, tuple) else item
            is_delete = str(value).strip().upper() != "KEEP"

            if is_delete and range_start is None:
                range_start = row_number
            elif not is_delete and range_start is not None:
                delete_ranges.append((range_start, row_number - 1))
                range_start = None

        if range_start is not None:
            delete_ranges.append((range_start, last_row))

        delete_count = sum(
            end_row - start_row + 1
            for start_row, end_row in delete_ranges
        )

        # Sanity check BEFORE deletion.
        expected_delete_count = total_rows - visible_count
        if delete_count != expected_delete_count:
            raise RuntimeError(
                "Current View marker validation failed before row deletion. "
                f"Excel marked {visible_count:,} KEEP row(s), so "
                f"{expected_delete_count:,} DELETE row(s) were expected, but "
                f"{delete_count:,} DELETE row(s) were mapped."
            )

        log(
            f"Current View deletion plan: {delete_count:,} row(s) "
            f"in {len(delete_ranges):,} contiguous block(s)."
        )

        # The visible state has already been captured into helper markers.
        # Clear filtering before physical row deletion.
        try:
            if ws.FilterMode:
                ws.ShowAllData()
        except Exception:
            pass

        try:
            if ws.AutoFilterMode:
                ws.AutoFilterMode = False
        except Exception:
            pass

        # Delete bottom-up so original row numbers remain valid.
        for index, (start_row, end_row) in enumerate(
            reversed(delete_ranges),
            start=1,
        ):
            try:
                self._call_with_com_retry(
                    f"deleting Current View rows {start_row}:{end_row}",
                    lambda start_row=start_row, end_row=end_row: ws.Rows(
                        f"{start_row}:{end_row}"
                    ).Delete(),
                    log=log,
                )
            except Exception as exc:
                raise RuntimeError(
                    "Excel could not delete rows outside the current visible view. "
                    f"Failed at rows {start_row}:{end_row} "
                    f"(block {index} of {len(delete_ranges)})."
                ) from exc

        # Validate using helper column, populated for every surviving row.
        remaining_rows = int(
            ws.Cells(ws.Rows.Count, helper_col).End(XL_UP).Row
        ) - header_row

        if remaining_rows != visible_count:
            raise RuntimeError(
                "Excel-native Current View validation failed after row deletion. "
                f"Expected {visible_count:,} visible row(s), but "
                f"{remaining_rows:,} row(s) remain."
            )

        self._call_with_com_retry(
            "clearing Current View helper cells",
            lambda: ws.Range(
                ws.Cells(header_row, helper_col),
                ws.Cells(last_row, helper_col),
            ).ClearContents(),
            log=log,
        )

        log(
            f"Current View conversion passed: "
            f"{visible_count:,} visible row(s), {delete_count:,} removed."
        )

        return visible_count, delete_count

    @staticmethod
    def _get_last_row(ws) -> int:
        return int(ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row)

    @staticmethod
    def _get_last_column(ws, header_row: int = 1) -> int:
        header_row = max(int(header_row), 1)
        header_last = int(
            ws.Cells(header_row, ws.Columns.Count).End(XL_TO_LEFT).Column
        )
        try:
            used = ws.UsedRange
            used_last = int(used.Column) + int(used.Columns.Count) - 1
            return max(header_last, used_last)
        except Exception:
            return header_last

    def _get_safe_helper_column(
        self,
        ws,
        minimum_column: int = 1,
        log: Optional[Callable[[str], None]] = None,
    ) -> int:
        """
        Return a temporary helper column that is safely to the right of workbook
        content and validation support ranges.

        The old implementation used the last non-blank HEADER cell + 1. That can
        collide with hidden dropdown-source columns whose header cells are blank.
        FIX5 deliberately considers UsedRange, Data Validation areas and Excel
        Names before choosing the helper column.
        """
        candidates = [max(int(minimum_column), 1), self._get_last_column(ws)]

        # UsedRange catches hidden support lists even when their header is blank.
        try:
            used = ws.UsedRange
            used_right = int(used.Column) + int(used.Columns.Count) - 1
            candidates.append(used_right)
        except Exception:
            pass

        # Validation target ranges may extend beyond the ordinary report area.
        try:
            validation_cells = ws.Cells.SpecialCells(
                XL_CELL_TYPE_ALL_VALIDATION
            )
            try:
                areas = list(validation_cells.Areas)
            except Exception:
                areas = [validation_cells]

            for area in areas:
                try:
                    right = int(area.Column) + int(area.Columns.Count) - 1
                    candidates.append(right)
                except Exception:
                    pass
        except Exception:
            pass

        # Workbook/worksheet Names can point to hidden helper ranges outside the
        # visible report table. RefersToRange is available only for range Names.
        for names_owner in (ws, getattr(ws, 'Parent', None)):
            if names_owner is None:
                continue
            try:
                names = names_owner.Names
            except Exception:
                continue

            try:
                count = int(names.Count)
            except Exception:
                count = 0

            for index in range(1, count + 1):
                try:
                    named_range = names.Item(index).RefersToRange
                    right = (
                        int(named_range.Column)
                        + int(named_range.Columns.Count)
                        - 1
                    )
                    candidates.append(right)
                except Exception:
                    # Constants / formulas / external names have no RefersToRange.
                    pass

        helper_col = max(candidates) + 1
        max_columns = int(ws.Columns.Count)

        if helper_col > max_columns:
            raise RuntimeError(
                'Excel helper-column safety check failed because there is no '
                'unused worksheet column available to the right of workbook data.'
            )

        if log:
            log(
                f'Temporary helper column selected safely at Excel column '
                f'{helper_col}; existing workbook/support columns end at '
                f'{helper_col - 1}.'
            )

        return helper_col

    def _apply_keep_rows_filter_fast(
        self,
        ws,
        keep_rows: set[int],
        log: Optional[Callable[[str], None]] = None,
        header_row: int = 1,
    ) -> None:
        """
        Fast reverse-filter deletion.

        A helper column marks KEEP/DELETE in one COM assignment. Excel filters
        DELETE rows, but deletion is requested through ONE reference column
        only. This avoids the large multi-column SpecialCells selection that
        becomes fragile when the source workbook has many hidden columns.

        The remaining row count is validated before the helper column is
        removed. A failed or incomplete deletion is never silently accepted.
        """
        try:
            if ws.FilterMode:
                ws.ShowAllData()
        except Exception:
            pass

        try:
            if ws.AutoFilterMode:
                ws.AutoFilterMode = False
        except Exception:
            pass

        header_row = max(int(header_row), 1)
        try:
            used = ws.UsedRange
            used_last_row = int(used.Row) + int(used.Rows.Count) - 1
        except Exception:
            used_last_row = header_row

        last_row = max(self._get_last_row(ws), used_last_row)
        last_col = self._get_last_column(ws, header_row=header_row)

        if last_row <= header_row:
            return

        valid_keep = {
            int(row)
            for row in keep_rows
            if header_row < int(row) <= last_row
        }

        helper_col = self._get_safe_helper_column(
            ws,
            minimum_column=last_col,
            log=log,
        )
        ws.Cells(header_row, helper_col).Value = "__GF_KEEP__"

        markers = [
            ("KEEP",) if row in valid_keep else ("DELETE",)
            for row in range(header_row + 1, last_row + 1)
        ]

        if markers:
            ws.Range(
                ws.Cells(header_row + 1, helper_col),
                ws.Cells(last_row, helper_col),
            ).Value = markers

        filter_range = ws.Range(
            ws.Cells(header_row, 1),
            ws.Cells(last_row, helper_col),
        )
        field = helper_col - int(filter_range.Column) + 1
        self._call_with_com_retry(
            "applying reverse filter",
            lambda: filter_range.AutoFilter(
                Field=field,
                Criteria1="DELETE",
            ),
            log=log,
        )

        total_data_rows = last_row - header_row
        delete_count = total_data_rows - len(valid_keep)

        if log:
            log(
                f"Reverse filter: keeping {len(valid_keep):,} row(s), "
                f"deleting {delete_count:,} row(s)."
            )

        if delete_count > 0:
            try:
                # Important: use one reference column only. Hidden columns no
                # longer fragment the SpecialCells selection.
                delete_cells = ws.Range(
                    ws.Cells(header_row + 1, 1),
                    ws.Cells(last_row, 1),
                ).SpecialCells(XL_CELL_TYPE_VISIBLE)

                self._call_with_com_retry(
                    "deleting filtered rows",
                    lambda: delete_cells.EntireRow.Delete(),
                    log=log,
                )

            except Exception as exc:
                raise RuntimeError(
                    "Excel could not delete the filtered rows. "
                    "The split was stopped to prevent creation of an "
                    "incorrect workbook."
                ) from exc

        try:
            if ws.AutoFilterMode:
                ws.AutoFilterMode = False
        except Exception:
            pass

        # The helper column has a value on every remaining data row, so it is
        # a reliable way to validate the row count even when column A contains
        # blanks.
        remaining_rows = int(
            ws.Cells(ws.Rows.Count, helper_col).End(XL_UP).Row
        ) - header_row

        expected_rows = len(valid_keep)

        if remaining_rows != expected_rows:
            raise RuntimeError(
                "Reverse-filter validation failed. "
                f"Expected {expected_rows:,} remaining row(s), but Excel "
                f"contains {remaining_rows:,}. The split was stopped before "
                "the workbook could be accepted as output."
            )

        self._call_with_com_retry(
            "clearing reverse-filter helper cells",
            lambda: ws.Range(
                ws.Cells(header_row, helper_col),
                ws.Cells(last_row, helper_col),
            ).ClearContents(),
            log=log,
        )

    def _apply_reverse_filter(
        self,
        ws,
        keep_rows: set[int],
        com_error,
        header_row: int = 1,
    ) -> None:
        # com_error is kept in the signature for compatibility with the
        # existing split engine. Errors are intentionally NOT swallowed.
        self._apply_keep_rows_filter_fast(
            ws,
            keep_rows,
            log=None,
            header_row=header_row,
        )

    @staticmethod
    def _contiguous_numbers(values: list[int]) -> list[tuple[int, int]]:
        if not values:
            return []

        values = sorted(set(int(v) for v in values))
        result: list[tuple[int, int]] = []
        start = previous = values[0]

        for value in values[1:]:
            if value == previous + 1:
                previous = value
                continue

            result.append((start, previous))
            start = previous = value

        result.append((start, previous))
        return result

    def _get_data_validation_info(
        self,
        ws,
        log: Optional[Callable[[str], None]] = None,
    ) -> dict:
        """Detect native Excel Data Validation without changing it."""
        result = {
            "has_validation": False,
            "has_list_validation": False,
        }

        try:
            cells = ws.Cells.SpecialCells(XL_CELL_TYPE_ALL_VALIDATION)
        except Exception:
            return result

        result["has_validation"] = True

        try:
            areas = list(cells.Areas)
        except Exception:
            areas = [cells]

        for area in areas:
            try:
                if int(area.Validation.Type) == XL_VALIDATE_LIST:
                    result["has_list_validation"] = True
                    break
            except Exception:
                continue

        if log:
            if result["has_list_validation"]:
                log(
                    "Excel List/dropdown Data Validation detected. "
                    "Native validation will be preserved."
                )
            elif result["has_validation"]:
                log("Excel Data Validation detected.")

        return result

    def _capture_data_validations(
        self,
        ws,
        keep_rows: set[int],
        log: Optional[Callable[[str], None]] = None,
    ) -> dict:
        """
        Capture Excel Data Validation rules before physical row deletion.

        Each validation Area is mapped from original worksheet row numbers to the
        compacted output row numbers created by the Split engine. The actual rule
        is restored after row deletion so dropdowns survive even when Excel drops
        validation while deleting large row blocks.
        """
        snapshot = {
            "rules": [],
            "has_validation": False,
            "has_list_validation": False,
        }

        try:
            validation_cells = ws.Cells.SpecialCells(
                XL_CELL_TYPE_ALL_VALIDATION
            )
        except Exception:
            return snapshot

        ordered_keep_rows = sorted(
            int(row)
            for row in keep_rows
            if int(row) >= 2
        )
        new_row_map = {
            original_row: 2 + index
            for index, original_row in enumerate(ordered_keep_rows)
        }

        try:
            areas = list(validation_cells.Areas)
        except Exception:
            areas = [validation_cells]

        for area in areas:
            try:
                row_start = int(area.Row)
                row_end = row_start + int(area.Rows.Count) - 1
                col_start = int(area.Column)
                col_end = col_start + int(area.Columns.Count) - 1

                validation = area.Validation
                validation_type = int(validation.Type)

                rule = {
                    "type": validation_type,
                    "alert_style": int(validation.AlertStyle),
                    "operator": int(validation.Operator),
                    "formula1": validation.Formula1,
                    "formula2": validation.Formula2,
                    "ignore_blank": bool(validation.IgnoreBlank),
                    "in_cell_dropdown": bool(validation.InCellDropdown),
                    "show_input": bool(validation.ShowInput),
                    "show_error": bool(validation.ShowError),
                    "input_title": str(validation.InputTitle or ""),
                    "error_title": str(validation.ErrorTitle or ""),
                    "input_message": str(validation.InputMessage or ""),
                    "error_message": str(validation.ErrorMessage or ""),
                    "ranges": [],
                }

                # Header/non-data validation: preserve its original row.
                if row_end < 2:
                    rule["ranges"].append(
                        (row_start, row_end, col_start, col_end)
                    )
                else:
                    surviving_original_rows = [
                        row
                        for row in ordered_keep_rows
                        if row_start <= row <= row_end
                    ]

                    mapped_rows = [
                        new_row_map[row]
                        for row in surviving_original_rows
                    ]

                    for new_start, new_end in self._contiguous_numbers(mapped_rows):
                        rule["ranges"].append(
                            (new_start, new_end, col_start, col_end)
                        )

                if rule["ranges"]:
                    snapshot["rules"].append(rule)
                    snapshot["has_validation"] = True

                    if validation_type == XL_VALIDATE_LIST:
                        snapshot["has_list_validation"] = True

            except Exception as exc:
                if log:
                    log(
                        "Warning: one Data Validation area could not be captured: "
                        f"{exc}"
                    )

        if log and snapshot["has_validation"]:
            list_text = (
                " including List/dropdown validation"
                if snapshot["has_list_validation"]
                else ""
            )
            log(
                f"Captured {len(snapshot['rules'])} Data Validation rule area(s)"
                f"{list_text} for restoration after row deletion."
            )

        return snapshot

    def _restore_data_validations(
        self,
        ws,
        snapshot: dict,
        log: Optional[Callable[[str], None]] = None,
    ) -> int:
        """
        Reapply captured validation rules to the compacted output rows.
        """
        rules = list(snapshot.get("rules", []))
        restored_ranges = 0

        for rule in rules:
            for row_start, row_end, col_start, col_end in rule["ranges"]:
                target = ws.Range(
                    ws.Cells(row_start, col_start),
                    ws.Cells(row_end, col_end),
                )

                try:
                    self._call_with_com_retry(
                        "clearing existing Data Validation before restore",
                        lambda target=target: target.Validation.Delete(),
                        log=log,
                    )
                except Exception:
                    # A range with no surviving validation can legitimately raise.
                    pass

                def add_validation(
                    target=target,
                    rule=rule,
                ):
                    kwargs = {
                        "Type": int(rule["type"]),
                        "AlertStyle": int(rule["alert_style"]),
                        "Operator": int(rule["operator"]),
                        "Formula1": rule["formula1"],
                    }

                    # Excel can reject Formula2 for validation types that do not use it.
                    formula2 = rule.get("formula2")
                    if formula2 not in (None, ""):
                        kwargs["Formula2"] = formula2

                    target.Validation.Add(**kwargs)

                self._call_with_com_retry(
                    "restoring Data Validation rule",
                    add_validation,
                    log=log,
                )

                validation = target.Validation

                for attr, value in (
                    ("IgnoreBlank", bool(rule["ignore_blank"])),
                    ("InCellDropdown", bool(rule["in_cell_dropdown"])),
                    ("ShowInput", bool(rule["show_input"])),
                    ("ShowError", bool(rule["show_error"])),
                    ("InputTitle", rule["input_title"]),
                    ("ErrorTitle", rule["error_title"]),
                    ("InputMessage", rule["input_message"]),
                    ("ErrorMessage", rule["error_message"]),
                ):
                    try:
                        setattr(validation, attr, value)
                    except Exception:
                        pass

                restored_ranges += 1

        if log and restored_ranges:
            log(
                f"Restored Data Validation to {restored_ranges} output range(s)."
            )

        return restored_ranges

    def _validate_dropdown_presence(
        self,
        ws,
        expected_list_validation: bool,
    ) -> None:
        """
        Safety check: if the original output copy contained list/dropdown
        validation, the finished output must still contain validation.
        """
        if not expected_list_validation:
            return

        try:
            cells = ws.Cells.SpecialCells(
                XL_CELL_TYPE_ALL_VALIDATION
            )
        except Exception as exc:
            raise RuntimeError(
                "Data Validation safety check failed. "
                "The source contained dropdown validation, but no validation "
                "cells remain in the split output."
            ) from exc

        found_list = False

        try:
            areas = list(cells.Areas)
        except Exception:
            areas = [cells]

        for area in areas:
            try:
                if int(area.Validation.Type) == XL_VALIDATE_LIST:
                    found_list = True
                    break
            except Exception:
                continue

        if not found_list:
            raise RuntimeError(
                "Data Validation safety check failed. "
                "List/dropdown validation was present before splitting but "
                "could not be verified in the output workbook."
            )

    def _delete_hidden_columns(
        self,
        ws,
        log: Optional[Callable[[str], None]] = None,
    ) -> int:
        used = ws.UsedRange
        start_col = int(used.Column)
        count = int(used.Columns.Count)
        hidden = []
        for col in range(start_col, start_col + count):
            try:
                if bool(ws.Columns(col).Hidden):
                    hidden.append(col)
            except Exception:
                pass
        for col in reversed(hidden):
            self._call_with_com_retry(
                f"deleting hidden column {col}",
                lambda col=col: ws.Columns(col).Delete(),
                log=log,
            )
        return len(hidden)

    def _make_visible_view_master(
        self,
        original_source: str,
        sheet_name: str,
        log: Callable[[str], None],
        worksheet_password: str = "",
        expected_analysis_rows: Optional[int] = None,
        header_row: int = 1,
    ) -> tuple[Path, int]:
        """
        Create a local visible-only master from the CURRENT Excel view.

        The current open source workbook is never unprotected, filtered, edited,
        saved, or closed. SaveCopyAs preserves its current Excel view. All edits
        happen only in the temporary copy.
        """
        pythoncom, win32, _ = _require_pywin32()
        pythoncom.CoInitialize()

        attached_excel = None
        source_wb = None
        temp_excel = None
        temp_wb = None
        temp_ws = None

        temp_file = (
            Path(tempfile.gettempdir())
            / "GoldenFirReportCenter"
            / "SplitReports"
            / f"VISIBLE_{datetime.now():%Y%m%d_%H%M%S_%f}_{Path(original_source).name}"
        )
        temp_file.parent.mkdir(parents=True, exist_ok=True)

        try:
            try:
                attached_excel = win32.GetObject(None, "Excel.Application")
            except Exception as exc:
                raise RuntimeError(
                    "Microsoft Excel is not open. Open the selected workbook first."
                ) from exc

            source_wb = self._find_open_workbook(attached_excel, original_source)
            if source_wb is None:
                raise RuntimeError(
                    "Please open the selected workbook in Excel first."
                )

            # Copy FIRST. The source workbook remains completely untouched.
            self._call_with_com_retry(
                "creating Current View SaveCopyAs",
                lambda: source_wb.SaveCopyAs(str(temp_file)),
                log=log,
            )
            log("Created temporary SaveCopyAs from the current open Excel workbook.")

            temp_excel = win32.DispatchEx("Excel.Application")
            temp_excel.Visible = False
            temp_excel.DisplayAlerts = False
            temp_excel.ScreenUpdating = False
            temp_excel.EnableEvents = False

            temp_wb = self._call_with_com_retry(
                "opening temporary Current View workbook",
                lambda: temp_excel.Workbooks.Open(
                    str(temp_file),
                    UpdateLinks=0,
                    ReadOnly=False,
                ),
                log=log,
            )
            temp_ws = temp_wb.Worksheets(sheet_name)

            protection_state = self._capture_protection_settings(temp_ws)

            if protection_state["is_protected"]:
                log(
                    "Temporary worksheet is protected; unlocking temporary copy "
                    "for Current View conversion."
                )
                self._unprotect_sheet(
                    temp_ws,
                    worksheet_password,
                    protection_state,
                )

            try:
                visible_count, _ = self._convert_current_view_in_place(
                    temp_ws,
                    log,
                    header_row=header_row,
                )
            finally:
                if protection_state["is_protected"]:
                    self._restore_sheet_protection(
                        temp_ws,
                        protection_state,
                        worksheet_password,
                    )
                    log("Worksheet protection restored on temporary view master.")

            self._call_with_com_retry(
                "saving temporary Current View workbook",
                lambda: temp_wb.Save(),
                log=log,
            )
            self._call_with_com_retry(
                "closing temporary Current View workbook",
                lambda: temp_wb.Close(SaveChanges=False),
                log=log,
            )
            temp_wb = None
            temp_ws = None

            # Re-read through pandas exactly as the split engine will.
            validation_df = pd.read_excel(
                temp_file,
                sheet_name=sheet_name,
                header=max(int(header_row), 1) - 1,
                dtype=object,
            )
            actual_rows = int(len(validation_df))

            log(f"Rows after Excel-native Current View conversion: {actual_rows:,}")

            # Excel-native validation above is authoritative for physical row
            # preservation/deletion. pandas/openpyxl can legitimately report fewer
            # rows when the visible Excel range ends with fully blank/formatted rows
            # (for example rows carried by AutoFilter/UsedRange formatting).
            #
            # A pandas count GREATER than Excel's visible count is still unsafe and
            # remains a hard failure because it implies unexpected rows survived.
            if actual_rows > visible_count:
                raise RuntimeError(
                    "Current View validation failed after saving the temporary "
                    f"workbook. Excel marked {visible_count:,} visible row(s), "
                    f"but pandas reads {actual_rows:,} row(s), which is more than "
                    "Excel's validated visible-row count."
                )

            if actual_rows < visible_count:
                blank_tail_rows = visible_count - actual_rows
                log(
                    "Current View note: Excel preserved "
                    f"{visible_count:,} visible row(s), while pandas reads "
                    f"{actual_rows:,} data row(s). The {blank_tail_rows:,} row "
                    "difference is accepted because Excel-native deletion validation "
                    "already passed; pandas/openpyxl can omit fully blank/formatted "
                    "rows at the end of the visible range."
                )

            if (
                expected_analysis_rows is not None
                and actual_rows != int(expected_analysis_rows)
            ):
                raise RuntimeError(
                    "Current Excel View changed after Analyze. "
                    f"Analyze found {int(expected_analysis_rows):,} row(s), but "
                    f"Start Splitting now prepares {actual_rows:,}. "
                    "Split cancelled before output creation. Run Analyze again."
                )

            log("Excel-native Current View validation passed.")
            return temp_file, actual_rows

        finally:
            try:
                if temp_wb is not None:
                    self._call_with_com_retry(
                        "closing temporary Current View workbook during cleanup",
                        lambda: temp_wb.Close(SaveChanges=False),
                        log=log,
                    )
            except Exception:
                pass

            try:
                if temp_excel is not None:
                    temp_excel.Quit()
            except Exception:
                pass

            temp_ws = None
            temp_wb = None
            temp_excel = None
            source_wb = None
            attached_excel = None
            gc.collect()

            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


    def _build_row_map(
        self,
        source_file: Path,
        sheet_name: str,
        split_column: str,
        header_row: int = 1,
    ):
        header_row = max(int(header_row), 1)
        df = pd.read_excel(
            source_file,
            sheet_name=sheet_name,
            header=header_row - 1,
            dtype=object,
        )
        if split_column not in df.columns:
            raise ValueError(f"Split column not found: {split_column}")
        row_map: dict[object, list[int]] = {}
        blank_rows = 0
        for idx, value in enumerate(
            df[split_column],
            start=header_row + 1,
        ):
            if not _normalise(value):
                blank_rows += 1
                continue
            row_map.setdefault(value, []).append(idx)
        groups = sorted(row_map.keys(), key=_normalise)
        return df, row_map, groups, blank_rows

    def _split_with_excel(
        self,
        source_file: Path,
        sheet_name: str,
        split_column: str,
        output_folder: Path,
        exclude_hidden_columns: bool,
        context,
        expected_rows: Optional[int] = None,
        filename_suffix: str = "",
        add_date_suffix: bool = False,
        worksheet_password: str = "",
        expected_groups: Optional[int] = None,
        header_row: int = 1,
        include_other_worksheets: bool = True,
    ) -> dict:
        pythoncom, win32, com_error = _require_pywin32()
        context.log("Reading source data...")
        df, row_map, groups, blank_rows = self._build_row_map(
            source_file,
            sheet_name,
            split_column,
            header_row=header_row,
        )
        context.log(f"Rows loaded: {len(df):,}")

        if expected_rows is not None and len(df) != expected_rows:
            raise RuntimeError(
                "Visible-view safety check failed before splitting. "
                f"Expected {expected_rows:,} row(s), but {len(df):,} row(s) "
                "were loaded. Split cancelled before creating output files."
            )

        context.log(f"Output groups found: {len(groups):,}")

        if expected_groups is not None and len(groups) != int(expected_groups):
            raise RuntimeError(
                "Analyze safety check failed before splitting. "
                f"Analyze found {int(expected_groups):,} output group(s), but "
                f"the prepared data contains {len(groups):,}. "
                "Split cancelled before creating output files. "
                "Run Analyze again with the current settings."
            )

        context.log(f"Blank split values skipped: {blank_rows:,}")
        if not groups:
            return {"files_created": 0, "files": [], "rows_loaded": len(df), "blank_rows": blank_rows}

        pythoncom.CoInitialize()
        excel = None
        wb = None
        created: list[str] = []
        protected_outputs = 0
        protection_restored = True
        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            total = len(groups)
            for index, group in enumerate(groups, start=1):
                filename = self.build_output_filename(
                    group,
                    custom_suffix=filename_suffix,
                    add_date_suffix=add_date_suffix,
                    extension=source_file.suffix,
                )
                output_file = output_folder / filename
                context.log(f"Creating {index}/{total}: {group}")
                shutil.copy2(source_file, output_file)

                if output_file.suffix.lower() != source_file.suffix.lower():
                    raise RuntimeError(
                        "Output workbook format safety check failed. "
                        f"Source format is {source_file.suffix}, but output format is "
                        f"{output_file.suffix}. Split cancelled before Excel opened "
                        "the copied workbook."
                    )

                wb = self._call_with_com_retry(
                    f"opening output workbook {filename}",
                    lambda: excel.Workbooks.Open(
                        str(output_file),
                        UpdateLinks=0,
                        ReadOnly=False,
                    ),
                    log=context.log,
                )
                ws = wb.Worksheets(sheet_name)
                protection_state = self._capture_protection_settings(ws)

                if protection_state["is_protected"]:
                    protected_outputs += 1
                    self._unprotect_sheet(
                        ws,
                        worksheet_password,
                        protection_state,
                    )
                    context.log(f"{filename}: worksheet temporarily unprotected.")

                try:
                    keep_rows = set(row_map[group])

                    validation_info = self._get_data_validation_info(
                        ws,
                        log=lambda message: context.log(
                            f"{filename}: {message}"
                        ),
                    )

                    self._apply_keep_rows_filter_fast(
                        ws,
                        keep_rows,
                        log=lambda message: context.log(
                            f"{filename}: {message}"
                        ),
                        header_row=header_row,
                    )

                    # Fidelity-first rule:
                    # if the workbook contains List/dropdown validation, never
                    # delete hidden columns. Those columns may contain the source
                    # list, Named Range backing data, or INDIRECT dependencies.
                    if exclude_hidden_columns:
                        if validation_info.get("has_list_validation", False):
                            context.log(
                                f"{filename}: Exclude Hidden Columns overridden. "
                                "Hidden columns retained because List/dropdown "
                                "Data Validation is present."
                            )
                        else:
                            removed = self._delete_hidden_columns(
                                ws,
                                log=lambda message: context.log(
                                    f"{filename}: {message}"
                                ),
                            )
                            if removed:
                                context.log(
                                    f"Excluded {removed} hidden column(s) from {filename}"
                                )

                    # Do not rebuild Validation through Validation.Add().
                    # The workbook is a native Excel copy, so preserve Excel's
                    # own validation objects and only verify that dropdowns remain.
                    self._validate_dropdown_presence(
                        ws,
                        bool(validation_info.get("has_list_validation", False)),
                    )

                    if validation_info.get("has_list_validation", False):
                        context.log(
                            f"{filename}: Native List/dropdown Data Validation verified."
                        )
                finally:
                    if protection_state["is_protected"]:
                        try:
                            self._restore_sheet_protection(
                                ws,
                                protection_state,
                                worksheet_password,
                            )
                            context.log(
                                f"{filename}: worksheet protection restored."
                            )
                        except Exception:
                            protection_restored = False
                            raise

                if not include_other_worksheets:
                    removed_sheets = 0
                    for sheet_index in range(int(wb.Worksheets.Count), 0, -1):
                        candidate = wb.Worksheets(sheet_index)
                        try:
                            candidate_name = str(candidate.Name)
                        except Exception:
                            candidate_name = ""

                        if candidate_name == sheet_name:
                            continue

                        self._call_with_com_retry(
                            f"deleting worksheet {candidate_name or sheet_index}",
                            lambda candidate=candidate: candidate.Delete(),
                            log=context.log,
                        )
                        removed_sheets += 1

                    context.log(
                        f"{filename}: removed {removed_sheets} other worksheet(s)."
                    )

                self._call_with_com_retry(
                    f"saving output workbook {filename}",
                    lambda: wb.Save(),
                    log=context.log,
                )
                self._call_with_com_retry(
                    f"closing output workbook {filename}",
                    lambda: wb.Close(SaveChanges=False),
                    log=context.log,
                )
                wb = None

                if protection_state["is_protected"]:
                    self._verify_saved_worksheet_password(
                        excel=excel,
                        output_file=output_file,
                        sheet_name=sheet_name,
                        protection_state=protection_state,
                        password=worksheet_password,
                        log=context.log,
                    )

                created.append(str(output_file))
                context.progress(int(index / total * 100), f"Creating {group} ({index}/{total})")
        finally:
            try:
                if wb is not None:
                    self._call_with_com_retry(
                        "closing output workbook during cleanup",
                        lambda: wb.Close(SaveChanges=False),
                        log=context.log,
                    )
            except Exception:
                pass
            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                pass
            wb = None
            excel = None
            gc.collect()
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        return {
            "files_created": len(created),
            "files": created,
            "rows_loaded": len(df),
            "blank_rows": blank_rows,
            "protected_outputs": protected_outputs,
            "protection_restored": protection_restored,
        }

    @staticmethod
    def _write_split_log(output_folder: Path, values: dict) -> None:
        path = output_folder / "Split_Log.txt"
        lines = ["=" * 48, "Golden Fir Report Center - Split Reports", "=" * 48, ""]
        for key, value in values.items():
            lines.extend([f"{key}:", str(value), ""])
        lines.append("=" * 48)
        path.write_text("\n".join(lines), encoding="utf-8")

    def run(
        self,
        source_file: str,
        sheet_name: str,
        split_column: str,
        base_output: str,
        exclude_hidden_columns: bool,
        use_open_excel_view: bool,
        filename_suffix: str,
        add_date_suffix: bool,
        worksheet_password: str,
        context,
        expected_analysis_rows: Optional[int] = None,
        expected_analysis_groups: Optional[int] = None,
        header_row: int = 1,
        include_other_worksheets: bool = True,
    ) -> JobResult:
        started = datetime.now()
        source = Path(source_file)
        if not source.is_file():
            raise FileNotFoundError(f"Source file not found: {source_file}")
        if not base_output:
            raise ValueError("Select an output folder first.")
        if not sheet_name:
            raise ValueError("Select a worksheet first.")
        if not split_column:
            raise ValueError("Select a split column first.")

        output = self._create_output_folder(base_output)
        context.log(f"Output folder: {output}")
        processing: Optional[Path] = None
        cleanup_processing = False
        expected_visible_rows: Optional[int] = None

        try:
            if use_open_excel_view:
                context.log("Using currently open Excel workbook view.")
                processing, expected_visible_rows = self._make_visible_view_master(
                    source_file,
                    sheet_name,
                    context.log,
                    worksheet_password=worksheet_password,
                    expected_analysis_rows=expected_analysis_rows,
                    header_row=header_row,
                )
                cleanup_processing = True
            else:
                processing, cleanup_processing = self._prepare_local_source(
                    source_file,
                    context.log,
                )

            summary = self._split_with_excel(
                processing,
                sheet_name,
                split_column,
                output,
                exclude_hidden_columns,
                context,
                expected_rows=expected_visible_rows,
                filename_suffix=filename_suffix,
                add_date_suffix=add_date_suffix,
                worksheet_password=worksheet_password,
                expected_groups=expected_analysis_groups,
                header_row=header_row,
                include_other_worksheets=include_other_worksheets,
            )
            duration = datetime.now() - started
            self._write_split_log(
                output,
                {
                    "Input File": source_file,
                    "Processing File": processing,
                    "Worksheet": sheet_name,
                    "Split By": split_column,
                    "Header Row": header_row,
                    "Include Other Worksheets": include_other_worksheets,
                    "Exclude Hidden Columns": exclude_hidden_columns,
                    "Data Mode": (
                        "Current Excel Workbook View"
                        if use_open_excel_view
                        else "Full Workbook Data"
                    ),
                    "Use Currently Open Excel Workbook View": use_open_excel_view,
                    "Analyze Expected Rows": (
                        expected_analysis_rows
                        if expected_analysis_rows is not None
                        else "(Not analyzed before split)"
                    ),
                    "Analyze Expected Output Groups": (
                        expected_analysis_groups
                        if expected_analysis_groups is not None
                        else "(Not analyzed before split)"
                    ),
                    "Filename Custom Suffix": filename_suffix or "(None)",
                    "Add Current Date Suffix": add_date_suffix,
                    "Filename Example": self.build_output_filename(
                        "Example Split Value",
                        custom_suffix=filename_suffix,
                        add_date_suffix=add_date_suffix,
                        extension=Path(processing).suffix if processing else Path(source_file).suffix,
                    ),
                    "Worksheet Protected": bool(summary.get("protected_outputs", 0)),
                    "Protected Output Files": summary.get("protected_outputs", 0),
                    "Worksheet Protection Restored": summary.get(
                        "protection_restored", True
                    ),
                    "Worksheet Password Persistence Verified": bool(
                        summary.get("protected_outputs", 0)
                    ),
                    "Rows Loaded": summary["rows_loaded"],
                    "Files Created": summary["files_created"],
                    "Blank Rows Skipped": summary["blank_rows"],
                    "Duration": duration,
                    "Status": "Completed Successfully",
                },
            )
            context.log("Split completed successfully.")
            context.log(f"Files created: {summary['files_created']}")
            context.log(f"Blank rows skipped: {summary['blank_rows']}")
            context.log(f"Output folder: {output}")
            return JobResult(
                True,
                f"Split completed: {summary['files_created']} file(s)",
                summary["files"],
                summary["files_created"],
                summary["blank_rows"],
                0,
            )
        except Exception as exc:
            try:
                self._write_split_log(
                    output,
                    {
                        "Input File": source_file,
                        "Worksheet": sheet_name,
                        "Split By": split_column,
                        "Header Row": header_row,
                        "Include Other Worksheets": include_other_worksheets,
                        "Error": str(exc),
                        "Status": "Failed",
                    },
                )
            except Exception:
                pass
            raise
        finally:
            if cleanup_processing and processing:
                try:
                    processing.unlink(missing_ok=True)
                except Exception:
                    pass
