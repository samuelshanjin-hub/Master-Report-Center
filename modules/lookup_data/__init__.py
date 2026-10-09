from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any
import shutil
import time

from core.models.job_result import JobResult


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb", ".xls"}
HEADER_ROW = 1
XL_UP = -4162
XL_TO_LEFT = -4159
XL_PASTE_FORMATS = -4122


@dataclass(frozen=True)
class LookupHeaderAnalysis:
    reference_headers: tuple[str, ...]
    target_headers: tuple[str, ...]
    common_headers: tuple[str, ...]
    return_headers: tuple[str, ...]
    reference_sheet: str
    target_sheet: str


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
        isinstance(value, str) and value.strip() == ""
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


def _worksheet(workbook: Any, requested_name: str):
    requested_name = requested_name.strip()

    if requested_name:
        try:
            return workbook.Worksheets(requested_name)
        except Exception as exc:
            raise RuntimeError(
                f"Worksheet '{requested_name}' was not found in "
                f"{Path(workbook.FullName).name}."
            ) from exc

    return workbook.Worksheets(1)


def _headers(ws: Any) -> tuple[list[str], dict[str, int]]:
    last_col = int(
        ws.Cells(
            HEADER_ROW,
            ws.Columns.Count,
        ).End(XL_TO_LEFT).Column
    )

    raw = ws.Range(
        ws.Cells(HEADER_ROW, 1),
        ws.Cells(HEADER_ROW, last_col),
    ).Value2

    if last_col == 1:
        values = [raw]
    else:
        values = list(raw[0])

    headers = [
        "" if value is None else str(value)
        for value in values
    ]

    normalised = [
        _normalise_header(value)
        for value in headers
    ]

    while normalised and normalised[-1] == "":
        normalised.pop()
        headers.pop()

    if not headers:
        raise RuntimeError("No headers were found in row 1.")

    if any(header == "" for header in normalised):
        raise RuntimeError(
            "Header row contains one or more blank headers."
        )

    if len(set(normalised)) != len(normalised):
        raise RuntimeError(
            "Header row contains duplicate header names."
        )

    return headers, {
        header: index + 1
        for index, header in enumerate(normalised)
    }


def _last_row_for_columns(
    ws: Any,
    columns: list[int],
) -> int:
    last_row = HEADER_ROW

    for col in columns:
        candidate = int(
            ws.Cells(
                ws.Rows.Count,
                col,
            ).End(XL_UP).Row
        )
        last_row = max(last_row, candidate)

    return last_row


def _read_column(
    ws: Any,
    col: int,
    start_row: int,
    end_row: int,
) -> list[Any]:
    if end_row < start_row:
        return []

    value = ws.Range(
        ws.Cells(start_row, col),
        ws.Cells(end_row, col),
    ).Value2

    count = end_row - start_row + 1

    if count == 1:
        return [value]

    return [
        row[0] if isinstance(row, (tuple, list)) else row
        for row in value
    ]


def _elapsed_text(seconds: float) -> str:
    return str(timedelta(seconds=seconds))


class LookupDataService:
    """
    Lookup / enrichment engine optimized for:
        large Reference Database -> smaller Target workbook.

    Design:
    - Build Target key requirements first.
    - Read only selected Match + Return columns from the Reference.
    - Keep only Reference rows whose keys are actually needed by Target.
    - Write values only into Target cells so Target formatting remains intact.
    - By default only fill blank Target cells.
    - Never append unrelated Reference records to Target.
    - Never overwrite the original Target workbook.
    """

    def analyze(
        self,
        reference_file: str,
        target_file: str,
        reference_sheet: str = "",
        target_sheet: str = "",
        context=None,
    ) -> LookupHeaderAnalysis:
        try:
            import pythoncom
            import win32com.client as win32
        except ImportError as exc:
            raise RuntimeError(
                "Lookup Data requires Microsoft Excel and pywin32 on Windows."
            ) from exc

        reference = Path(reference_file).expanduser().resolve()
        target = Path(target_file).expanduser().resolve()

        if not reference.is_file():
            raise FileNotFoundError(
                f"Reference database not found: {reference}"
            )

        if not target.is_file():
            raise FileNotFoundError(
                f"Target workbook not found: {target}"
            )

        pythoncom.CoInitialize()

        excel = reference_wb = target_wb = None

        def log(message: str) -> None:
            if context is not None:
                context.log(message)

        try:
            log("Starting Microsoft Excel for header analysis...")

            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False

            reference_wb = _call_with_com_retry(
                "opening Reference workbook for header analysis",
                lambda: excel.Workbooks.Open(
                    str(reference),
                    UpdateLinks=0,
                    ReadOnly=True,
                    AddToMru=False,
                    IgnoreReadOnlyRecommended=True,
                ),
                log=log,
            )

            target_wb = _call_with_com_retry(
                "opening Target workbook for header analysis",
                lambda: excel.Workbooks.Open(
                    str(target),
                    UpdateLinks=0,
                    ReadOnly=True,
                    AddToMru=False,
                    IgnoreReadOnlyRecommended=True,
                ),
                log=log,
            )

            reference_ws = _worksheet(
                reference_wb,
                reference_sheet,
            )
            target_ws = _worksheet(
                target_wb,
                target_sheet,
            )

            reference_headers, _ = _headers(reference_ws)
            target_headers, _ = _headers(target_ws)

            reference_by_norm = {
                _normalise_header(header): header
                for header in reference_headers
            }

            target_by_norm = {
                _normalise_header(header): header
                for header in target_headers
            }

            common = tuple(
                target_by_norm[norm]
                for norm in target_by_norm
                if norm in reference_by_norm
            )

            # Return fields are driven by the Reference database.
            # Match fields will be excluded in the page once the user selects them.
            return_headers = tuple(reference_headers)

            log(f"Reference sheet: {reference_ws.Name}")
            log(f"Target sheet: {target_ws.Name}")
            log(f"Reference headers: {len(reference_headers)}")
            log(f"Target headers: {len(target_headers)}")
            log(f"Common Match candidates: {len(common)}")

            if context is not None:
                context.progress(
                    100,
                    "Header analysis completed.",
                )

            return LookupHeaderAnalysis(
                reference_headers=tuple(reference_headers),
                target_headers=tuple(target_headers),
                common_headers=common,
                return_headers=return_headers,
                reference_sheet=str(reference_ws.Name),
                target_sheet=str(target_ws.Name),
            )

        finally:
            _safe_close(reference_wb)
            _safe_close(target_wb)

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
        reference_file: Path,
        target_file: Path,
        reference_sheet: str,
        target_sheet: str,
        match_headers: tuple[str, ...],
        return_headers: tuple[str, ...],
        fill_blank_only: bool,
        ignore_blank_reference: bool,
        reference_rows_read: int,
        target_rows_read: int,
        matched_target_rows: int,
        unmatched_target_rows: int,
        cells_written: int,
        existing_cells_preserved: int,
        blank_reference_values_skipped: int,
        duplicate_reference_keys: int,
        missing_target_fields_added: int,
        output_file: Path | None,
        duration_seconds: float,
        status: str,
        error_message: str | None = None,
    ) -> None:
        lines = [
            "================================================",
            "Golden Fir Report Center - Lookup Data",
            "================================================",
            "",
            "Reference Database:",
            str(reference_file),
            "",
            "Target Workbook:",
            str(target_file),
            "",
            "Reference Worksheet:",
            reference_sheet.strip() or "(First worksheet)",
            "",
            "Target Worksheet:",
            target_sheet.strip() or "(First worksheet)",
            "",
            "Match Fields:",
            " + ".join(match_headers),
            "",
            "Return Fields:",
            ", ".join(return_headers),
            "",
            "Only Fill Blank Target Cells:",
            str(bool(fill_blank_only)),
            "",
            "Ignore Blank Reference Values:",
            str(bool(ignore_blank_reference)),
            "",
            "Reference Rows Read:",
            str(reference_rows_read),
            "",
            "Target Rows Read:",
            str(target_rows_read),
            "",
            "Matched Target Rows:",
            str(matched_target_rows),
            "",
            "Unmatched Target Rows:",
            str(unmatched_target_rows),
            "",
            "Cells Written:",
            str(cells_written),
            "",
            "Existing Target Cells Preserved:",
            str(existing_cells_preserved),
            "",
            "Blank Reference Values Skipped:",
            str(blank_reference_values_skipped),
            "",
            "Duplicate Reference Keys:",
            str(duplicate_reference_keys),
            "",
            "Missing Target Fields Added:",
            str(missing_target_fields_added),
            "",
            "Output File:",
            str(output_file) if output_file else "(None)",
            "",
            "Duration:",
            _elapsed_text(duration_seconds),
            "",
            "Status:",
            status,
        ]

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

        log_file.parent.mkdir(parents=True, exist_ok=True)
        log_file.write_text(
            "\n".join(lines),
            encoding="utf-8",
        )

    def run(
        self,
        reference_file: str,
        target_file: str,
        reference_sheet: str,
        target_sheet: str,
        match_headers: tuple[str, ...],
        return_headers: tuple[str, ...],
        fill_blank_only: bool,
        ignore_blank_reference: bool,
        context,
    ) -> JobResult:
        try:
            import pythoncom
            import win32com.client as win32
        except ImportError as exc:
            raise RuntimeError(
                "Lookup Data requires Microsoft Excel and pywin32 on Windows."
            ) from exc

        started = time.perf_counter()

        reference = Path(reference_file).expanduser().resolve()
        target = Path(target_file).expanduser().resolve()

        if not reference.is_file():
            raise FileNotFoundError(
                f"Reference database not found: {reference}"
            )

        if not target.is_file():
            raise FileNotFoundError(
                f"Target workbook not found: {target}"
            )

        if reference.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise RuntimeError(
                f"Unsupported Reference file type: {reference.suffix}"
            )

        if target.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise RuntimeError(
                f"Unsupported Target file type: {target.suffix}"
            )

        if not match_headers:
            raise RuntimeError(
                "Select at least one Match field."
            )

        if not return_headers:
            raise RuntimeError(
                "Select at least one Return field."
            )

        overlap = sorted(
            set(match_headers) & set(return_headers)
        )

        if overlap:
            raise RuntimeError(
                "Match fields cannot also be selected as Return fields: "
                + ", ".join(overlap)
            )

        timestamp = datetime.now().strftime(
            "%Y-%m-%d_%H%M%S"
        )

        output_file = target.with_name(
            f"{target.stem}_Lookup_"
            f"{timestamp}{target.suffix}"
        )

        log_file = target.with_name(
            f"Lookup_Log_{timestamp}.txt"
        )

        shutil.copy2(
            target,
            output_file,
        )

        reference_rows_read = 0
        target_rows_read = 0
        matched_target_rows = 0
        unmatched_target_rows = 0
        cells_written = 0
        existing_cells_preserved = 0
        blank_reference_values_skipped = 0
        duplicate_reference_keys = 0
        missing_target_fields_added = 0

        def report(percent: int, message: str) -> None:
            safe_percent = max(
                0,
                min(100, int(percent)),
            )
            context.progress(
                safe_percent,
                message,
            )
            context.log(message)

        pythoncom.CoInitialize()

        excel = reference_wb = output_wb = None

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
                report(
                    2,
                    "Excel calculation mode set to manual.",
                )
            except Exception:
                report(
                    2,
                    "Excel calculation mode could not be changed; continuing normally.",
                )

            # ----------------------------------------------------------
            # Phase 1 â€” Read Target keys first.
            # ----------------------------------------------------------
            report(
                5,
                "Reading Target workbook and building lookup requirements...",
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

            target_ws = _worksheet(
                output_wb,
                target_sheet,
            )

            target_headers, target_map = _headers(target_ws)

            match_norm = tuple(
                _normalise_header(header)
                for header in match_headers
            )

            return_norm = tuple(
                _normalise_header(header)
                for header in return_headers
            )

            for header, original in zip(
                match_norm,
                match_headers,
            ):
                if header not in target_map:
                    raise RuntimeError(
                        f"Match field not found in Target: {original}"
                    )

            target_last_row = _last_row_for_columns(
                target_ws,
                [
                    target_map[header]
                    for header in match_norm
                ],
            )

            if target_last_row < HEADER_ROW + 1:
                raise RuntimeError(
                    "Target workbook contains no data rows."
                )

            target_count = target_last_row - HEADER_ROW

            target_match_columns = {
                header: _read_column(
                    target_ws,
                    target_map[header],
                    HEADER_ROW + 1,
                    target_last_row,
                )
                for header in match_norm
            }

            target_keys: list[
                tuple[tuple[str, Any], ...] | None
            ] = []

            needed_keys: set[
                tuple[tuple[str, Any], ...]
            ] = set()

            for index in range(target_count):
                key = tuple(
                    _normalise_key_value(
                        target_match_columns[header][index]
                    )
                    for header in match_norm
                )

                target_rows_read += 1

                if all(
                    component[0] == "blank"
                    for component in key
                ):
                    target_keys.append(None)
                    continue

                target_keys.append(key)
                needed_keys.add(key)

            report(
                15,
                (
                    f"Target rows read: {target_rows_read:,}; "
                    f"unique lookup keys required: {len(needed_keys):,}."
                ),
            )

            # ----------------------------------------------------------
            # Phase 2 â€” Read only selected columns from Reference.
            # ----------------------------------------------------------
            report(
                20,
                "Opening Reference database...",
            )

            reference_wb = _call_with_com_retry(
                f"opening Reference workbook {reference.name}",
                lambda: excel.Workbooks.Open(
                    str(reference),
                    UpdateLinks=0,
                    ReadOnly=True,
                    AddToMru=False,
                    IgnoreReadOnlyRecommended=True,
                ),
                log=context.log,
            )

            reference_ws = _worksheet(
                reference_wb,
                reference_sheet,
            )

            reference_headers, reference_map = _headers(
                reference_ws
            )

            for header, original in zip(
                match_norm,
                match_headers,
            ):
                if header not in reference_map:
                    raise RuntimeError(
                        f"Match field not found in Reference: {original}"
                    )

            for header, original in zip(
                return_norm,
                return_headers,
            ):
                if header not in reference_map:
                    raise RuntimeError(
                        f"Return field not found in Reference: {original}"
                    )

            required_reference_columns = sorted(
                {
                    reference_map[header]
                    for header in (
                        match_norm + return_norm
                    )
                }
            )

            reference_last_row = _last_row_for_columns(
                reference_ws,
                required_reference_columns,
            )

            if reference_last_row < HEADER_ROW + 1:
                raise RuntimeError(
                    "Reference database contains no data rows."
                )

            reference_count = (
                reference_last_row - HEADER_ROW
            )

            report(
                25,
                (
                    f"Reference rows to scan: {reference_count:,}. "
                    "Reading only selected Match and Return columns..."
                ),
            )

            reference_match_columns = {
                header: _read_column(
                    reference_ws,
                    reference_map[header],
                    HEADER_ROW + 1,
                    reference_last_row,
                )
                for header in match_norm
            }

            reference_return_columns = {
                header: _read_column(
                    reference_ws,
                    reference_map[header],
                    HEADER_ROW + 1,
                    reference_last_row,
                )
                for header in return_norm
            }

            # key -> tuple(return values), first Reference occurrence wins.
            lookup_index: dict[
                tuple[tuple[str, Any], ...],
                tuple[Any, ...],
            ] = {}

            for index in range(reference_count):
                key = tuple(
                    _normalise_key_value(
                        reference_match_columns[header][index]
                    )
                    for header in match_norm
                )

                reference_rows_read += 1

                if all(
                    component[0] == "blank"
                    for component in key
                ):
                    continue

                # Optimization: only retain rows needed by the Target.
                if key not in needed_keys:
                    continue

                if key in lookup_index:
                    duplicate_reference_keys += 1
                    continue

                lookup_index[key] = tuple(
                    reference_return_columns[header][index]
                    for header in return_norm
                )

                if (
                    len(lookup_index) == len(needed_keys)
                    and reference_rows_read < reference_count
                ):
                    # Every Target key has already been resolved.
                    # We can stop scanning the in-memory column arrays early.
                    break

                if (
                    reference_rows_read == 1
                    or reference_rows_read % 10000 == 0
                ):
                    progress = 25 + int(
                        min(
                            reference_rows_read,
                            reference_count,
                        )
                        / max(1, reference_count)
                        * 35
                    )
                    report(
                        progress,
                        (
                            f"Scanning Reference: "
                            f"{reference_rows_read:,} / {reference_count:,} rows..."
                        ),
                    )

            report(
                62,
                (
                    f"Reference lookup index built: "
                    f"{len(lookup_index):,} required keys found; "
                    f"{duplicate_reference_keys:,} duplicate key occurrence(s) ignored."
                ),
            )

            # ----------------------------------------------------------
            # Phase 3 â€” Ensure Return fields exist in Target.
            # ----------------------------------------------------------
            current_target_headers, target_map = _headers(
                target_ws
            )

            last_target_col = len(
                current_target_headers
            )

            for original_header, normalised_header in zip(
                return_headers,
                return_norm,
            ):
                if normalised_header in target_map:
                    continue

                last_target_col += 1

                new_header_cell = target_ws.Cells(
                    HEADER_ROW,
                    last_target_col,
                )

                # Copy format from previous header where possible.
                if last_target_col > 1:
                    try:
                        _call_with_com_retry(
                            "copying new Target header format",
                            lambda: target_ws.Cells(
                                HEADER_ROW,
                                last_target_col - 1,
                            ).Copy(),
                            log=context.log,
                        )
                        _call_with_com_retry(
                            "pasting new Target header format",
                            lambda: new_header_cell.PasteSpecial(
                                XL_PASTE_FORMATS
                            ),
                            log=context.log,
                        )
                    except Exception:
                        pass

                _call_with_com_retry(
                    f"writing new Target header {original_header}",
                    lambda original_header=original_header: setattr(
                        new_header_cell,
                        "Value2",
                        original_header,
                    ),
                    log=context.log,
                )
                target_map[normalised_header] = last_target_col
                missing_target_fields_added += 1

            # ----------------------------------------------------------
            # Phase 4 â€” Fill / update Target.
            # ----------------------------------------------------------
            report(
                68,
                "Writing lookup results into Target copy...",
            )

            matched_row_flags = [False] * target_count

            for row_offset, key in enumerate(
                target_keys,
                start=0,
            ):
                if key is None:
                    unmatched_target_rows += 1
                    continue

                result = lookup_index.get(key)

                if result is None:
                    unmatched_target_rows += 1
                    continue

                matched_target_rows += 1
                matched_row_flags[row_offset] = True

                excel_row = HEADER_ROW + 1 + row_offset

                for field_index, (
                    normalised_header,
                    reference_value,
                ) in enumerate(
                    zip(
                        return_norm,
                        result,
                    )
                ):
                    if (
                        ignore_blank_reference
                        and _is_blank(reference_value)
                    ):
                        blank_reference_values_skipped += 1
                        continue

                    target_cell = target_ws.Cells(
                        excel_row,
                        target_map[normalised_header],
                    )

                    existing_value = target_cell.Value2

                    if (
                        fill_blank_only
                        and not _is_blank(existing_value)
                    ):
                        existing_cells_preserved += 1
                        continue

                    # Values Only by design:
                    # preserve Target workbook formatting / validation / layout.
                    _call_with_com_retry(
                        "writing Lookup value to Target",
                        lambda reference_value=reference_value: setattr(
                            target_cell,
                            "Value2",
                            reference_value,
                        ),
                        log=context.log,
                    )
                    cells_written += 1

                if (
                    matched_target_rows == 1
                    or matched_target_rows % 100 == 0
                    or (
                        matched_target_rows
                        + unmatched_target_rows
                        == target_count
                    )
                ):
                    progress = 68 + int(
                        (
                            matched_target_rows
                            + unmatched_target_rows
                        )
                        / max(1, target_count)
                        * 25
                    )
                    report(
                        progress,
                        (
                            f"Processing Target row "
                            f"{matched_target_rows + unmatched_target_rows:,} "
                            f"of {target_count:,}..."
                        ),
                    )

            report(
                95,
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

            _safe_close(
                reference_wb,
                save_changes=False,
            )
            reference_wb = None

            elapsed = time.perf_counter() - started

            self._write_log(
                log_file,
                reference_file=reference,
                target_file=target,
                reference_sheet=reference_sheet,
                target_sheet=target_sheet,
                match_headers=match_headers,
                return_headers=return_headers,
                fill_blank_only=fill_blank_only,
                ignore_blank_reference=ignore_blank_reference,
                reference_rows_read=reference_rows_read,
                target_rows_read=target_rows_read,
                matched_target_rows=matched_target_rows,
                unmatched_target_rows=unmatched_target_rows,
                cells_written=cells_written,
                existing_cells_preserved=existing_cells_preserved,
                blank_reference_values_skipped=blank_reference_values_skipped,
                duplicate_reference_keys=duplicate_reference_keys,
                missing_target_fields_added=missing_target_fields_added,
                output_file=output_file,
                duration_seconds=elapsed,
                status="Completed Successfully",
            )

            context.log("")
            context.log(
                "========== Lookup Summary =========="
            )
            context.log(
                f"Reference rows scanned: {reference_rows_read:,}"
            )
            context.log(
                f"Target rows read: {target_rows_read:,}"
            )
            context.log(
                f"Matched Target rows: {matched_target_rows:,}"
            )
            context.log(
                f"Unmatched Target rows: {unmatched_target_rows:,}"
            )
            context.log(
                f"Cells written: {cells_written:,}"
            )
            context.log(
                "Existing Target cells preserved: "
                f"{existing_cells_preserved:,}"
            )
            context.log(
                "Blank Reference values skipped: "
                f"{blank_reference_values_skipped:,}"
            )
            context.log(
                "Duplicate Reference keys ignored: "
                f"{duplicate_reference_keys:,}"
            )
            context.log(
                "Missing Target fields added: "
                f"{missing_target_fields_added:,}"
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
                "Lookup completed.",
            )

            return JobResult(
                success=True,
                message=(
                    f"Lookup completed; "
                    f"{matched_target_rows} Target row(s) matched and "
                    f"{cells_written} cell(s) written."
                ),
                output_files=[
                    str(output_file),
                    str(log_file),
                ],
                processed_count=target_rows_read,
                skipped_count=unmatched_target_rows,
                warning_count=(
                    duplicate_reference_keys
                    + unmatched_target_rows
                ),
            )

        except Exception as exc:
            elapsed = time.perf_counter() - started

            try:
                self._write_log(
                    log_file,
                    reference_file=reference,
                    target_file=target,
                    reference_sheet=reference_sheet,
                    target_sheet=target_sheet,
                    match_headers=match_headers,
                    return_headers=return_headers,
                    fill_blank_only=fill_blank_only,
                    ignore_blank_reference=ignore_blank_reference,
                    reference_rows_read=reference_rows_read,
                    target_rows_read=target_rows_read,
                    matched_target_rows=matched_target_rows,
                    unmatched_target_rows=unmatched_target_rows,
                    cells_written=cells_written,
                    existing_cells_preserved=existing_cells_preserved,
                    blank_reference_values_skipped=blank_reference_values_skipped,
                    duplicate_reference_keys=duplicate_reference_keys,
                    missing_target_fields_added=missing_target_fields_added,
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
            _safe_close(reference_wb)
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

