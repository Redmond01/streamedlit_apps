"""Core engine for Report Auto-Sync on Station Daily Reports.

Automates the rollover of station daily report workbooks:
1. Copies the latest daily sheet (index -1) to create a new sheet (new index -1).
2. Sets Date = previous sheet date + 1 day.
3. Resets Outer & Inner Meters: Closing -> Opening, Closing = Opening (Sales = 0).
4. Resets Tank Dipping: Closing -> Opening, Closing = Opening (Dipping Sales = 0).
5. Clears POS list and updates rollover date.
6. Clears Products Received.
7. Clears Cash-to-bank, Expenses, and Imprest.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
import io
import logging
from pathlib import Path
import re
from typing import Any, Callable
import zipfile

import openpyxl

from core.memory import cleanup_files, force_gc, spool_uploaded_file

logger = logging.getLogger(__name__)


@dataclass
class SyncSummary:
    filename: str
    station_name: str
    old_sheet: str
    new_sheet: str
    old_date: str
    new_date: str
    outer_pumps_updated: int
    inner_pumps_updated: int
    tanks_updated: int
    pos_cleared: int
    products_cleared: int
    expenses_cleared: int
    status: str = "Success"
    message: str = "Sheet successfully created and synchronized."


def _parse_cell_date(cell_val: Any) -> datetime | None:
    """Parses a date or datetime object from a cell value."""
    if isinstance(cell_val, datetime):
        return cell_val
    if isinstance(cell_val, date):
        return datetime.combine(cell_val, datetime.min.time())
    if cell_val is None:
        return None
    val_str = str(cell_val).strip()
    if not val_str:
        return None
    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%d/%m/%Y %H:%M:%S",
        "%d-%m-%Y",
        "%d-%m-%Y %H:%M:%S",
        "%m/%d/%Y",
    ):
        try:
            return datetime.strptime(val_str, fmt)
        except ValueError:
            pass
    return None


def _format_next_sheet_title(src_title: str) -> str:
    """Determines the next sheet title from the source sheet title (e.g. Sheet10 -> Sheet11)."""
    m = re.search(r"^(.*?)(\d+)$", src_title.strip())
    if m:
        base, num = m.group(1), int(m.group(2))
        return f"{base}{num + 1}"
    m_space = re.search(r"^(.*?)\s+(\d+)$", src_title.strip())
    if m_space:
        base, num = m_space.group(1), int(m_space.group(2))
        return f"{base} {num + 1}"
    return f"{src_title}_new"


def sync_station_report(
    wb: openpyxl.Workbook,
) -> tuple[openpyxl.Workbook, SyncSummary]:
    """Synchronizes a station workbook by copying the last sheet and resetting the 7 sectors.

    Args:
        wb: openpyxl Workbook loaded with data_only=False.

    Returns:
        tuple of (updated_wb, SyncSummary)
    """
    if not wb.worksheets:
        raise ValueError("Workbook contains no worksheets.")

    src_ws = wb.worksheets[-1]
    new_ws = wb.copy_worksheet(src_ws)

    # 1. Determine new sheet title
    target_title = _format_next_sheet_title(src_ws.title)
    existing_titles = {ws.title.lower() for ws in wb.worksheets if ws != new_ws}
    candidate = target_title
    counter = 1
    while candidate.lower() in existing_titles:
        candidate = f"{target_title}_{counter}"
        counter += 1
    new_ws.title = candidate

    # 2. Date: Locate date in row 1 (A1:F1) and increment by 1 day
    old_date_dt = None
    new_date_dt = None
    date_cell = None
    for c in range(1, min(10, new_ws.max_column + 1)):
        val = new_ws.cell(1, c).value
        parsed = _parse_cell_date(val)
        if parsed:
            old_date_dt = parsed
            date_cell = new_ws.cell(1, c)
            break

    if old_date_dt:
        new_date_dt = old_date_dt + timedelta(days=1)
        if isinstance(date_cell.value, str):
            if " " in date_cell.value:
                date_cell.value = new_date_dt.strftime("%Y-%m-%d %H:%M:%S")
            else:
                date_cell.value = new_date_dt.strftime("%Y-%m-%d")
        else:
            date_cell.value = new_date_dt

    # 3. Locate Tank Table Header Row First (to establish meter row boundary)
    tank_header_row = None
    outer_tank_open_col = None
    outer_tank_close_col = None
    inner_tank_open_col = None
    inner_tank_close_col = None

    for r in range(10, min(30, new_ws.max_row + 1)):
        for c in range(1, new_ws.max_column + 1):
            val = str(new_ws.cell(r, c).value or "").strip().upper()
            if "OPENNG DIPPG" in val or "OPENING DIPPG" in val:
                if outer_tank_open_col is None:
                    outer_tank_open_col = c
                    tank_header_row = r
                else:
                    inner_tank_open_col = c
            elif "CLOSNG DIPPG" in val or "CLOSING DIPPG" in val:
                if outer_tank_close_col is None:
                    outer_tank_close_col = c
                else:
                    inner_tank_close_col = c

    # 4. Locate Meter Table Headers (Outer and Inner)
    meter_header_row = None
    outer_open_col = None
    inner_open_col = None
    outer_close_cols: list[int] = []
    inner_close_cols: list[int] = []

    for r in range(1, 6):
        for c in range(1, new_ws.max_column + 1):
            val = str(new_ws.cell(r, c).value or "").strip().upper()
            if "OPENING METER" in val:
                if outer_open_col is None:
                    outer_open_col = c
                    meter_header_row = r
                else:
                    inner_open_col = c

    if meter_header_row:
        for c in range(1, new_ws.max_column + 1):
            val = str(new_ws.cell(meter_header_row, c).value or "").strip().upper()
            if "CLOSING METER" in val:
                if inner_open_col and c > inner_open_col:
                    inner_close_cols.append(c)
                else:
                    outer_close_cols.append(c)

    # 5. Reset Outer & Inner Meters: Closing -> Opening, Closing = Opening (Sales = 0)
    outer_pumps_updated = 0
    inner_pumps_updated = 0

    meter_scan_limit = tank_header_row if tank_header_row else min(25, new_ws.max_row + 1)

    if meter_header_row and outer_open_col:
        for r in range(meter_header_row + 1, meter_scan_limit):
            val = src_ws.cell(r, outer_open_col).value
            lbl = str(src_ws.cell(r, outer_open_col - 1).value or "").strip().upper()
            lbl_prev = str(src_ws.cell(r, outer_open_col - 2).value or "").strip().upper() if outer_open_col >= 3 else ""

            is_pump = (
                (isinstance(val, (int, float)) or (isinstance(val, str) and val.replace(".", "").isdigit()))
                and "TOTAL" not in lbl
                and "TOTAL" not in lbl_prev
            )

            if is_pump:
                # Find final active closing meter from src_ws
                last_closing = None
                for col_c in outer_close_cols:
                    cv = src_ws.cell(r, col_c).value
                    if cv is not None and isinstance(cv, (int, float)):
                        last_closing = cv
                    elif cv is not None and isinstance(cv, str) and cv.replace(".", "").isdigit():
                        last_closing = float(cv) if "." in cv else int(cv)

                if last_closing is not None:
                    new_ws.cell(r, outer_open_col).value = last_closing
                    # Set active closing meter columns to match opening (sales = 0)
                    for col_c in outer_close_cols:
                        if src_ws.cell(r, col_c).value is not None:
                            new_ws.cell(r, col_c).value = last_closing
                        else:
                            new_ws.cell(r, col_c).value = None
                    outer_pumps_updated += 1

            # Inner Meter Table reset
            if inner_open_col:
                in_val = src_ws.cell(r, inner_open_col).value
                is_inner_pump = (
                    (isinstance(in_val, (int, float)) or (isinstance(in_val, str) and in_val.replace(".", "").isdigit()))
                    and "TOTAL" not in lbl
                    and "TOTAL" not in lbl_prev
                )
                if is_inner_pump:
                    last_in_closing = None
                    for col_c in inner_close_cols:
                        cv = src_ws.cell(r, col_c).value
                        if cv is not None and isinstance(cv, (int, float)):
                            last_in_closing = cv
                        elif cv is not None and isinstance(cv, str) and cv.replace(".", "").isdigit():
                            last_in_closing = float(cv) if "." in cv else int(cv)

                    if last_in_closing is not None:
                        new_ws.cell(r, inner_open_col).value = last_in_closing
                        for col_c in inner_close_cols:
                            if src_ws.cell(r, col_c).value is not None:
                                new_ws.cell(r, col_c).value = last_in_closing
                            else:
                                new_ws.cell(r, col_c).value = None
                        inner_pumps_updated += 1

    # 6. Reset Tank Dippings: Closing -> Opening, Closing = Opening (Dipping Sales = 0)
    tanks_updated = 0
    if tank_header_row and outer_tank_open_col and outer_tank_close_col:
        for r in range(tank_header_row + 1, min(tank_header_row + 15, new_ws.max_row + 1)):
            lbl = str(new_ws.cell(r, outer_tank_open_col - 1).value or "").strip().upper()
            val = src_ws.cell(r, outer_tank_open_col).value

            # Stop when encountering an empty row, TOTAL row, or POS section
            if val is None and not lbl:
                break
            if "TOTAL" in lbl or "ADD UP" in lbl or "ZENITH" in lbl or "OPAY" in lbl:
                break

            if isinstance(val, (int, float)) or (isinstance(val, str) and val.replace(".", "").isdigit()):
                last_dip = src_ws.cell(r, outer_tank_close_col).value
                if last_dip is not None:
                    new_ws.cell(r, outer_tank_open_col).value = last_dip
                    new_ws.cell(r, outer_tank_close_col).value = last_dip

                if inner_tank_open_col and inner_tank_close_col:
                    in_last_dip = src_ws.cell(r, inner_tank_close_col).value
                    if in_last_dip is not None:
                        new_ws.cell(r, inner_tank_open_col).value = in_last_dip
                        new_ws.cell(r, inner_tank_close_col).value = in_last_dip
                tanks_updated += 1

    # 7. Products Received: Clear Litres to Empty
    products_cleared = 0
    for r in range(1, min(25, new_ws.max_row + 1)):
        for c in range(1, new_ws.max_column + 1):
            val = str(new_ws.cell(r, c).value or "").strip().upper()
            if "PRODUCTS RECEIVED" in val or "PRODUCT RECEIVED" in val:
                for r_sub in range(r + 1, min(r + 12, new_ws.max_row + 1)):
                    p_lbl = new_ws.cell(r_sub, c).value
                    p_lbl_str = str(p_lbl or "").strip().upper()
                    if p_lbl_str and "TOTAL" not in p_lbl_str and "EXPENSES" not in p_lbl_str:
                        target_cell = new_ws.cell(r_sub, c + 1)
                        if target_cell.value is not None:
                            target_cell.value = None
                            products_cleared += 1
                    elif "TOTAL" in p_lbl_str or "EXPENSES" in p_lbl_str:
                        break

    # 8. Expenses & Imprest: Clear manual expense entries
    expenses_cleared = 0
    for r in range(5, min(28, new_ws.max_row + 1)):
        for c in range(1, new_ws.max_column + 1):
            val = str(new_ws.cell(r, c).value or "").strip().upper()
            if val == "EXPENSES":
                for r_sub in range(r + 1, min(r + 15, new_ws.max_row + 1)):
                    amt_cell = new_ws.cell(r_sub, c + 1)
                    lbl_cell = new_ws.cell(r_sub, c)
                    lbl_str = str(lbl_cell.value or "").strip().upper()
                    if "TOTAL" in lbl_str:
                        break
                    if amt_cell.value is not None and not str(amt_cell.value).startswith("="):
                        amt_cell.value = None
                        expenses_cleared += 1
            elif "IMPREST" in val:
                for c_scan in (c + 1, c + 2):
                    v_imp = new_ws.cell(r, c_scan).value
                    if v_imp is not None and not str(v_imp).startswith("="):
                        new_ws.cell(r, c_scan).value = None
                        expenses_cleared += 1
                        break

    # 9. POS Table: Clear Values, Rollover Date, and Carry over Available Cash
    pos_cleared = 0
    for r in range(15, new_ws.max_row + 1):
        for c in range(1, min(30, new_ws.max_column + 1)):
            val = str(new_ws.cell(r, c).value or "").strip().upper()
            if "ADD UP" in val:
                header_row_1 = r
                val_row_1 = r + 1
                col_addup = c
                col_avail_cash = c + 1

                # Left side POS columns (cols 1 to c - 1)
                for c_pos in range(1, col_addup):
                    hdr_cell = new_ws.cell(header_row_1, c_pos).value
                    if hdr_cell is not None and not str(hdr_cell).startswith("="):
                        cell_to_clear = new_ws.cell(val_row_1, c_pos)
                        if cell_to_clear.value is not None and not str(cell_to_clear.value).startswith("="):
                            cell_to_clear.value = None
                            pos_cleared += 1

                # The ADD UP figure is static and must be brought forward without change (never an equation)
                addup_val = src_ws.cell(val_row_1, col_addup).value
                if isinstance(addup_val, str) and addup_val.startswith("="):
                    # Do not set equation into ADD UP
                    pass
                else:
                    new_ws.cell(val_row_1, col_addup).value = addup_val

                # Rollover row (Day N+1)
                if header_row_1 + 3 <= new_ws.max_row:
                    header_row_2 = header_row_1 + 2
                    val_row_2 = header_row_1 + 3

                    # Update rollover date cell (in Col A or B)
                    for c_d in (1, 2):
                        v_d = new_ws.cell(val_row_2, c_d).value
                        parsed_d = _parse_cell_date(v_d)
                        if parsed_d and new_date_dt:
                            rollover_dt = new_date_dt + timedelta(days=1)
                            if isinstance(v_d, str):
                                if " " in v_d:
                                    new_ws.cell(val_row_2, c_d).value = rollover_dt.strftime("%Y-%m-%d %H:%M:%S")
                                else:
                                    new_ws.cell(val_row_2, c_d).value = rollover_dt.strftime("%Y-%m-%d")
                            else:
                                new_ws.cell(val_row_2, c_d).value = rollover_dt

                    # Clear rollover POS values
                    for c_pos in range(1, col_addup):
                        hdr_cell = new_ws.cell(header_row_2, c_pos).value
                        if hdr_cell is not None and not str(hdr_cell).startswith("="):
                            cell_to_clear = new_ws.cell(val_row_2, c_pos)
                            if cell_to_clear.value is not None and not str(cell_to_clear.value).startswith("="):
                                cell_to_clear.value = None
                                pos_cleared += 1

                # Right side POS columns (from col_avail_cash + 1 to right ADD UP)
                right_addup_col = None
                for c_r in range(col_avail_cash + 1, new_ws.max_column + 1):
                    v_r = str(new_ws.cell(header_row_1, c_r).value or "").strip().upper()
                    if "ADD UP" in v_r:
                        right_addup_col = c_r
                        break

                end_right_pos = right_addup_col if right_addup_col else new_ws.max_column + 1
                for c_r in range(col_avail_cash + 1, end_right_pos):
                    hdr_cell = new_ws.cell(header_row_1, c_r).value
                    hdr_str = str(hdr_cell or "").strip().upper()
                    if "ADD UP" in hdr_str or "AVAILABLE" in hdr_str:
                        continue
                    if hdr_cell is not None and not str(hdr_cell).startswith("="):
                        cell_to_clear = new_ws.cell(val_row_1, c_r)
                        if cell_to_clear.value is not None and not str(cell_to_clear.value).startswith("="):
                            cell_to_clear.value = None
                            pos_cleared += 1

                # Ensure right-side ADD UP static figures are preserved
                for c_scan in range(col_avail_cash + 1, new_ws.max_column + 1):
                    if "ADD UP" in str(new_ws.cell(header_row_1, c_scan).value or "").upper():
                        new_ws.cell(val_row_1, c_scan).value = src_ws.cell(val_row_1, c_scan).value

                # Break after handling the primary POS table
                break

    # Station name inference
    sheet_sample = src_ws.title
    old_date_str = old_date_dt.strftime("%Y-%m-%d") if old_date_dt else "Unknown"
    new_date_str = new_date_dt.strftime("%Y-%m-%d") if new_date_dt else "Unknown"

    summary = SyncSummary(
        filename="",
        station_name="",
        old_sheet=src_ws.title,
        new_sheet=new_ws.title,
        old_date=old_date_str,
        new_date=new_date_str,
        outer_pumps_updated=outer_pumps_updated,
        inner_pumps_updated=inner_pumps_updated,
        tanks_updated=tanks_updated,
        pos_cleared=pos_cleared,
        products_cleared=products_cleared,
        expenses_cleared=expenses_cleared,
        status="Success",
        message=f"Created {new_ws.title} ({new_date_str}) from {src_ws.title} ({old_date_str}) with all meters, tanks, POS, and expenses reset.",
    )

    return wb, summary


def sync_station_file(
    input_file: Any,
) -> tuple[bytes, str, SyncSummary]:
    """Processes a single station file (UploadedFile or path) and returns the updated bytes, filename, and summary.

    Returns:
        tuple of (updated_excel_bytes, output_filename, SyncSummary)
    """
    if isinstance(input_file, (str, Path)):
        orig_name = Path(input_file).name
    else:
        orig_name = getattr(input_file, "name", "station_report.xlsx")
    spooled_path = spool_uploaded_file(input_file, prefix="report_sync_")

    try:
        wb = openpyxl.load_workbook(spooled_path, data_only=False)
        updated_wb, summary = sync_station_report(wb)

        # Infer station name from filename
        stem = Path(orig_name).stem
        summary.filename = orig_name
        summary.station_name = stem

        buf = io.BytesIO()
        updated_wb.save(buf)
        buf.seek(0)
        output_bytes = buf.getvalue()

        # Generate clean output filename
        out_name = f"{stem}_synced_{summary.new_date.replace('-', '')}.xlsx"
        return output_bytes, out_name, summary

    finally:
        cleanup_files(spooled_path)
        force_gc()


def sync_multiple_station_files(
    files: list[Any],
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[list[tuple[str, bytes]], io.BytesIO, list[SyncSummary]]:
    """Batch synchronizes multiple station workbooks and packages them into a single zip archive.

    Args:
        files: List of Streamlit UploadedFiles or paths.
        progress_callback: Optional callable(index, total, filename) for progress reporting.

    Returns:
        tuple of (individual_files_list, zip_buffer, summaries_list)
    """
    results: list[tuple[str, bytes]] = []
    summaries: list[SyncSummary] = []
    total = len(files)

    for idx, f in enumerate(files):
        name = getattr(f, "name", f"file_{idx+1}.xlsx")
        if progress_callback:
            progress_callback(idx, total, name)

        try:
            out_bytes, out_name, summary = sync_station_file(f)
            results.append((out_name, out_bytes))
            summaries.append(summary)
        except Exception as e:
            logger.error("Failed to sync %s: %s", name, e, exc_info=True)
            summaries.append(
                SyncSummary(
                    filename=name,
                    station_name=Path(name).stem,
                    old_sheet="Error",
                    new_sheet="Error",
                    old_date="Error",
                    new_date="Error",
                    outer_pumps_updated=0,
                    inner_pumps_updated=0,
                    tanks_updated=0,
                    pos_cleared=0,
                    products_cleared=0,
                    expenses_cleared=0,
                    status="Error",
                    message=str(e),
                )
            )

    # Package into a zip archive ensuring unique entry names
    zip_buf = io.BytesIO()
    used_names: set[str] = set()
    with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for out_name, data in results:
            unique_name = out_name
            c_dup = 1
            while unique_name in used_names:
                p = Path(out_name)
                unique_name = f"{p.stem}_{c_dup}{p.suffix}"
                c_dup += 1
            used_names.add(unique_name)
            zf.writestr(unique_name, data)

    zip_buf.seek(0)
    return results, zip_buf, summaries
