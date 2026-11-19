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


def _resolve_last_active_closing(
    src_ws: openpyxl.worksheet.worksheet.Worksheet | list[openpyxl.worksheet.worksheet.Worksheet],
    r: int,
    open_col: int,
    close_cols: list[int],
) -> tuple[int | float, bool]:
    """Finds the last valid active closing meter reading for a pump row strictly from src_ws (sheet -1).

    Prioritizes the latest positive non-zero reading from the closing columns of src_ws,
    ignoring trailing inactive placeholder zeros (common for single/double shift AGO/LPG).
    If no positive closing is found, falls back to src_ws opening meter.
    If src_ws only has zeros (e.g. decommissioned pump), strictly returns (0, False).
    Never looks back across older sheets (e.g. Sheet 4).

    Returns:
        tuple of (last_closing_value, any_positive_close_found)
    """
    if isinstance(src_ws, list):
        ws = src_ws[-1]
    else:
        ws = src_ws

    last_closing = None

    # 1. Search src_ws closing cols for the last positive number (> 0)
    for c in reversed(close_cols):
        if c <= ws.max_column:
            v = ws.cell(r, c).value
            if v is not None and isinstance(v, (int, float)) and v > 0:
                last_closing = v
                break
            if v is not None and isinstance(v, str) and v.replace(".", "").isdigit() and float(v) > 0:
                last_closing = float(v) if "." in v else int(v)
                break

    any_positive_close = last_closing is not None

    # 2. Check src_ws opening meter if positive
    if last_closing is None and open_col <= ws.max_column:
        v_op = ws.cell(r, open_col).value
        if v_op is not None and isinstance(v_op, (int, float)) and v_op > 0:
            last_closing = v_op
        elif v_op is not None and isinstance(v_op, str) and v_op.replace(".", "").isdigit() and float(v_op) > 0:
            last_closing = float(v_op) if "." in v_op else int(v_op)

    if last_closing is None:
        last_closing = 0

    return last_closing, any_positive_close


def _get_raw_or_data_num(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    r: int,
    c: int,
    ws_data: openpyxl.worksheet.worksheet.Worksheet | None = None,
) -> float:
    """Helper to get a numeric float from ws_data or ws."""
    if ws_data and c <= ws_data.max_column and r <= ws_data.max_row:
        v = ws_data.cell(r, c).value
        if isinstance(v, (int, float)):
            return float(v)
    if c <= ws.max_column and r <= ws.max_row:
        v = ws.cell(r, c).value
        if isinstance(v, (int, float)):
            return float(v)
        if isinstance(v, str):
            if v.startswith("="):
                return _eval_sub_cell_formula(ws, v, ws_data)
            try:
                return float(v.replace(",", "").strip())
            except ValueError:
                pass
    return 0.0


def _eval_sub_cell_formula(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    formula_str: str,
    ws_data: openpyxl.worksheet.worksheet.Worksheet | None = None,
) -> float:
    """Evaluates sub-formulas like '=SUM(S3:S15)' or '=SUM(V16:V22)'."""
    if not formula_str or not str(formula_str).startswith("="):
        try:
            return float(str(formula_str).replace(",", "").strip())
        except (ValueError, TypeError):
            return 0.0

    s = str(formula_str).strip()[1:].strip()
    if s.upper().startswith("SUM(") and s.endswith(")"):
        inner = s[4:-1].strip()
        if ":" in inner:
            parts = inner.split(":")
            m1 = re.match(r"^([A-Za-z]+)(\d+)$", parts[0].strip())
            m2 = re.match(r"^([A-Za-z]+)(\d+)$", parts[1].strip())
            if m1 and m2:
                c1 = openpyxl.utils.column_index_from_string(m1.group(1).upper())
                r1 = int(m1.group(2))
                c2 = openpyxl.utils.column_index_from_string(m2.group(1).upper())
                r2 = int(m2.group(2))
                tot = 0.0
                for r_i in range(r1, r2 + 1):
                    for c_i in range(c1, c2 + 1):
                        tot += _get_raw_or_data_num(ws, r_i, c_i, ws_data)
                return tot

    def _replace_ref(m):
        c_i = openpyxl.utils.column_index_from_string(m.group(1).upper())
        r_i = int(m.group(2))
        return str(_get_raw_or_data_num(ws, r_i, c_i, ws_data))

    expr = re.sub(r"([A-Za-z]+)(\d+)", _replace_ref, s)
    try:
        if re.match(r"^[0-9\.\+\-\*\/\(\)\s]+$", expr):
            return float(eval(expr))
    except Exception:
        pass
    return 0.0


def _evaluate_cash_formula(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    formula_str: str,
    ws_data: openpyxl.worksheet.worksheet.Worksheet | None = None,
) -> float | None:
    """Evaluates an Excel cash balance formula like '=SUM(S17-V23-C26-...-Q28+R26)' safely."""
    if not formula_str or not str(formula_str).startswith("="):
        return None

    s = str(formula_str).strip()[1:].strip()
    if s.upper().startswith("SUM(") and s.endswith(")"):
        s = s[4:-1].strip()

    s = s.replace("--", "+").replace("+-", "-").replace("-+", "-")
    tokens = re.findall(r"([+-]?\s*[A-Za-z0-9_]+)", s)
    total = 0.0
    found_any = False

    for tok in tokens:
        tok = tok.strip()
        if not tok:
            continue
        sign = -1.0 if tok.startswith("-") else 1.0
        tok = tok.lstrip("+-").strip()

        m = re.match(r"^([A-Za-z]+)(\d+)$", tok)
        if m:
            col_letter, row_num = m.group(1).upper(), int(m.group(2))
            c_idx = openpyxl.utils.column_index_from_string(col_letter)
            val = None
            if ws_data and c_idx <= ws_data.max_column and row_num <= ws_data.max_row:
                v = ws_data.cell(row_num, c_idx).value
                if isinstance(v, (int, float)):
                    val = float(v)
            if val is None and c_idx <= ws.max_column and row_num <= ws.max_row:
                v = ws.cell(row_num, c_idx).value
                if isinstance(v, (int, float)):
                    val = float(v)
                elif isinstance(v, str):
                    if not v.startswith("="):
                        try:
                            val = float(v.replace(",", "").strip())
                        except ValueError:
                            val = 0.0
                    else:
                        val = _eval_sub_cell_formula(ws, v, ws_data)
            if val is not None:
                total += sign * val
                found_any = True
        else:
            try:
                total += sign * float(tok.replace(",", "").strip())
                found_any = True
            except ValueError:
                pass

    return round(total, 2) if found_any else None


def _get_available_cash_value(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    val_row: int,
    col_avail: int | None,
    ws_data: openpyxl.worksheet.worksheet.Worksheet | None = None,
) -> float | None:
    """Extracts or evaluates the ending Available Cash from a sheet.

    Tries in priority:
    1. Pre-calculated float from ws_data (data_only=True).
    2. Static float from ws (if not a formula).
    3. Formula evaluation from ws/ws_data.
    """
    if col_avail is None or col_avail > ws.max_column or val_row > ws.max_row:
        return None

    # 1. Check data_only workbook first
    if ws_data and col_avail <= ws_data.max_column and val_row <= ws_data.max_row:
        v_d = ws_data.cell(val_row, col_avail).value
        if isinstance(v_d, (int, float)):
            return float(v_d)
        elif isinstance(v_d, str) and not v_d.startswith("="):
            try:
                return float(v_d.replace(",", "").strip())
            except ValueError:
                pass

    # 2. Check standard workbook
    v = ws.cell(val_row, col_avail).value
    if isinstance(v, (int, float)):
        return float(v)
    elif isinstance(v, str) and not v.startswith("="):
        try:
            return float(v.replace(",", "").strip())
        except ValueError:
            pass
    elif isinstance(v, str) and v.startswith("="):
        # 3. Evaluate formula if cached value was missing
        evaluated = _evaluate_cash_formula(ws, v, ws_data)
        if evaluated is not None:
            return evaluated

    return None


def clean_report_name(name_or_path: str | Path | Any) -> str:
    """Extracts a clean station report name from a filename, path, or UploadedFile.

    Strips intermediate suffixes like '_sync', '_synced', '_filled', '_reset',
    as well as appended timestamps/dates (e.g. '_20261117' or '_20261117_143640').
    Works iteratively to unwrap repeated suffixes cleanly.

    Examples:
        'NOVEMBER_Mowe_sgr_sync.xlsx' -> 'NOVEMBER_Mowe_sgr'
        'NOVEMBER_Mowe_sgr_synced_20261117.xlsx' -> 'NOVEMBER_Mowe_sgr'
        'NOVEMBER_Mowe_sgr_filled_20261117_143640.xlsx' -> 'NOVEMBER_Mowe_sgr'
        'NOVEMBER Jah michael Ta_sync.xlsx' -> 'NOVEMBER Jah michael Ta'
    """
    if not name_or_path:
        return "station_report"

    if hasattr(name_or_path, "name") and not isinstance(name_or_path, (str, Path)):
        name_str = getattr(name_or_path, "name", "station_report")
    else:
        name_str = str(name_or_path)

    stem = Path(name_str).stem
    pattern = re.compile(
        r"([-_ ]*(?:sync(?:ed)?|filled|reset))([-_ ]*\d{8}(?:[-_ ]*\d{4,6})?)?",
        re.IGNORECASE,
    )
    while True:
        new_stem = pattern.sub("", stem).strip(" _-")
        if new_stem == stem or not new_stem:
            break
        stem = new_stem

    return stem or Path(name_str).stem


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
    wb_data: openpyxl.Workbook | None = None,
) -> tuple[openpyxl.Workbook, SyncSummary]:
    """Synchronizes a station workbook by copying the last sheet and resetting the 7 sectors.

    Args:
        wb: openpyxl Workbook loaded with data_only=False.
        wb_data: Optional openpyxl Workbook loaded with data_only=True to extract evaluated figures.

    Returns:
        tuple of (updated_wb, SyncSummary)
    """
    if not wb.worksheets:
        raise ValueError("Workbook contains no worksheets.")

    src_ws = wb.worksheets[-1]
    new_ws = wb.copy_worksheet(src_ws)

    src_ws_data = None
    if wb_data is not None:
        if src_ws.title in wb_data.sheetnames:
            src_ws_data = wb_data[src_ws.title]
        elif wb_data.worksheets:
            src_ws_data = wb_data.worksheets[-1]

    # 1. Determine new sheet title
    target_title = _format_next_sheet_title(src_ws.title)
    existing_titles = {ws.title.lower() for ws in wb.worksheets if ws != new_ws}
    candidate = target_title
    counter = 1
    while candidate.lower() in existing_titles:
        candidate = f"{target_title}_{counter}"
        counter += 1
    new_ws.title = candidate

    # 2. Date: Locate date in row 1-3 (A1:J3) and increment by 1 day
    old_date_dt = None
    new_date_dt = None
    date_cell = None
    for r in range(1, 4):
        for c in range(1, min(10, new_ws.max_column + 1)):
            val = new_ws.cell(r, c).value
            parsed = _parse_cell_date(val)
            if parsed:
                old_date_dt = parsed
                date_cell = new_ws.cell(r, c)
                break
        if old_date_dt:
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

    # 3. Locate Tank Table Header Row First (to establish meter row boundary) - Multi-pair dynamic tanks
    tank_header_row = None
    dip_pairs: list[tuple[int, int]] = []

    for r in range(8, min(45, new_ws.max_row + 1)):
        row_pairs = []
        for c in range(1, new_ws.max_column):
            v1 = str(new_ws.cell(r, c).value or "").strip().upper()
            v2 = str(new_ws.cell(r, c + 1).value or "").strip().upper()
            is_open = "OPEN" in v1 and any(k in v1 for k in ["DIP", "DIPP", "DIPG"])
            is_close = any(k in v2 for k in ["CLOS", "CLOSE"]) and any(k in v2 for k in ["DIP", "DIPP", "DIPG"])
            if is_open and is_close:
                row_pairs.append((c, c + 1))
        if row_pairs:
            tank_header_row = r
            dip_pairs = row_pairs
            break

    # Fallback tank header detection if specific dipping keywords differ
    if tank_header_row is None:
        for r in range(8, min(45, new_ws.max_row + 1)):
            has_tank = any("TANK" in str(new_ws.cell(r, c).value or "").upper() for c in range(1, new_ws.max_column + 1))
            if has_tank:
                row_pairs = []
                for c in range(1, new_ws.max_column):
                    v1 = str(new_ws.cell(r, c).value or "").strip().upper()
                    v2 = str(new_ws.cell(r, c + 1).value or "").strip().upper()
                    if "OPEN" in v1 and ("CLOS" in v2 or "CLOSE" in v2):
                        row_pairs.append((c, c + 1))
                if row_pairs:
                    tank_header_row = r
                    dip_pairs = row_pairs
                    break

    outer_dip_pairs: list[tuple[int, int]] = []
    inner_dip_pairs: list[tuple[int, int]] = []
    if dip_pairs:
        first_c = dip_pairs[0][0]
        # Outer pairs are within 10 columns of the first pair
        outer_dip_pairs = [p for p in dip_pairs if p[0] - first_c < 10]
        # Inner pairs are separated by at least 10 columns from first pair
        inner_dip_pairs = [p for p in dip_pairs if p[0] - first_c >= 10]

    # For backward compatibility if any legacy code checks these:
    outer_tank_open_col = outer_dip_pairs[0][0] if outer_dip_pairs else None
    outer_tank_close_col = outer_dip_pairs[0][1] if outer_dip_pairs else None
    inner_tank_open_col = inner_dip_pairs[0][0] if inner_dip_pairs else None
    inner_tank_close_col = inner_dip_pairs[0][1] if inner_dip_pairs else None

    # 4. Locate Meter Table Headers (Outer and Inner)
    meter_header_row = None
    outer_open_col = None
    inner_open_col = None
    outer_close_cols: list[int] = []
    inner_close_cols: list[int] = []

    for r in range(1, 6):
        for c in range(1, new_ws.max_column + 1):
            val = str(new_ws.cell(r, c).value or "").strip().upper()
            if "OPENING METER" in val or "OPEN METER" in val:
                if outer_open_col is None:
                    outer_open_col = c
                    meter_header_row = r
                elif c > outer_open_col + 5:
                    inner_open_col = c

    # Fallback meter header detection if "OPENING METER" keyword is missing or mislabeled
    # E.g. in workbooks like Ogbomoso 2, the opening meter cell is labeled "CLOSING METER"
    if outer_open_col is None:
        for r in range(1, 6):
            meter_cols = []
            for c in range(1, new_ws.max_column + 1):
                v = str(new_ws.cell(r, c).value or "").strip().upper()
                if "METER" in v:
                    meter_cols.append(c)
            if meter_cols:
                outer_m = [c for c in meter_cols if c < 25]
                inner_m = [c for c in meter_cols if c >= 25]
                if outer_m:
                    outer_open_col = outer_m[0]
                    meter_header_row = r
                    outer_close_cols = outer_m[1:]
                if inner_m:
                    inner_open_col = inner_m[0]
                    inner_close_cols = inner_m[1:]
                if outer_open_col:
                    break

    # Secondary fallback: analyze sales formulas in row 3 or 4 (e.g. J3 = '=F3-E3')
    if outer_open_col is None:
        for r in (3, 4):
            if outer_open_col:
                break
            for c in range(1, min(25, new_ws.max_column + 1)):
                val = str(new_ws.cell(r, c).value or "").strip().upper()
                m = re.match(r"^=([A-Z]+)\d+-([A-Z]+)\d+$", val)
                if m:
                    close_col_letter, open_col_letter = m.group(1), m.group(2)
                    c_close = openpyxl.utils.column_index_from_string(close_col_letter)
                    c_open = openpyxl.utils.column_index_from_string(open_col_letter)
                    if c_open < c_close:
                        outer_open_col = c_open
                        meter_header_row = r - 1
                        outer_close_cols = [c_close]
                        break

    if inner_open_col is None and meter_header_row and new_ws.max_column > 25:
        for r in (3, 4):
            if inner_open_col:
                break
            for c in range(25, new_ws.max_column + 1):
                val = str(new_ws.cell(r, c).value or "").strip().upper()
                m = re.match(r"^=([A-Z]+)\d+-([A-Z]+)\d+$", val)
                if m:
                    close_col_letter, open_col_letter = m.group(1), m.group(2)
                    c_close = openpyxl.utils.column_index_from_string(close_col_letter)
                    c_open = openpyxl.utils.column_index_from_string(open_col_letter)
                    if c_open < c_close:
                        inner_open_col = c_open
                        inner_close_cols = [c_close]
                        break

    if meter_header_row and not outer_close_cols:
        for c in range(1, new_ws.max_column + 1):
            val = str(new_ws.cell(meter_header_row, c).value or "").strip().upper()
            if "CLOSING METER" in val or "CLOSE METER" in val:
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
            # Safeguard: stop immediately if any cell in the row indicates TANK table header
            row_vals_sample = [str(new_ws.cell(r, c).value or "").strip().upper() for c in range(1, min(10, new_ws.max_column + 1))]
            if any(v == "TANK" or v.startswith("TANK") for v in row_vals_sample):
                break

            val = src_ws.cell(r, outer_open_col).value
            lbl = str(src_ws.cell(r, outer_open_col - 1).value or "").strip().upper()

            # Outer pump: must have numeric opening meter and label must not be a subtotal/formula
            is_numeric_meter = (isinstance(val, (int, float)) or (isinstance(val, str) and val.replace(".", "").isdigit()))
            is_not_total = "TOTAL" not in lbl and not lbl.startswith("=")
            is_pump = is_numeric_meter and is_not_total

            if is_pump:
                # Find final active closing meter strictly from src_ws (sheet -1)
                last_closing, any_pos = _resolve_last_active_closing(
                    src_ws, r, outer_open_col, outer_close_cols
                )

                if last_closing is not None:
                    new_ws.cell(r, outer_open_col).value = last_closing
                    # Set active closing meter columns to match opening (sales = 0)
                    for col_c in outer_close_cols:
                        cv = src_ws.cell(r, col_c).value
                        if any_pos:
                            if cv is not None and (
                                (isinstance(cv, (int, float)) and cv > 0)
                                or (isinstance(cv, str) and cv.replace(".", "").isdigit() and float(cv) > 0)
                            ):
                                new_ws.cell(r, col_c).value = last_closing
                            elif cv == 0:
                                new_ws.cell(r, col_c).value = 0
                            else:
                                new_ws.cell(r, col_c).value = None
                        else:
                            if cv is not None:
                                new_ws.cell(r, col_c).value = last_closing
                            else:
                                new_ws.cell(r, col_c).value = None
                    outer_pumps_updated += 1

            # Inner Meter Table reset (evaluated independently with its own inner label and value)
            if inner_open_col:
                in_val = src_ws.cell(r, inner_open_col).value
                in_lbl = str(src_ws.cell(r, inner_open_col - 1).value or "").strip().upper()
                is_in_numeric = (
                    (isinstance(in_val, (int, float)) or (isinstance(in_val, str) and in_val.replace(".", "").isdigit()))
                    or any(isinstance(src_ws.cell(r, c).value, (int, float)) for c in inner_close_cols)
                )
                is_in_not_total = "TOTAL" not in in_lbl and not in_lbl.startswith("=")
                is_inner_pump = is_in_numeric and is_in_not_total

                if is_inner_pump:
                    last_in_closing, in_any_pos = _resolve_last_active_closing(
                        src_ws, r, inner_open_col, inner_close_cols
                    )

                    if last_in_closing is not None:
                        new_ws.cell(r, inner_open_col).value = last_in_closing
                        for col_c in inner_close_cols:
                            cv = src_ws.cell(r, col_c).value
                            if in_any_pos:
                                if cv is not None and (
                                    (isinstance(cv, (int, float)) and cv > 0)
                                    or (isinstance(cv, str) and cv.replace(".", "").isdigit() and float(cv) > 0)
                                ):
                                    new_ws.cell(r, col_c).value = last_in_closing
                                elif cv == 0:
                                    new_ws.cell(r, col_c).value = 0
                                else:
                                    new_ws.cell(r, col_c).value = None
                            else:
                                if cv is not None:
                                    new_ws.cell(r, col_c).value = last_in_closing
                                else:
                                    new_ws.cell(r, col_c).value = None
                        inner_pumps_updated += 1

    # 6. Reset Tank Dippings: Closing -> Opening, Closing = Opening (Dipping Sales = 0)
    # Dynamic across all physical, manifolded/formulated, and standalone tank pairs
    tanks_updated = 0
    if tank_header_row and outer_dip_pairs:
        first_open_col = outer_dip_pairs[0][0]
        for r in range(tank_header_row + 1, min(tank_header_row + 15, new_ws.max_row + 1)):
            lbl = str(new_ws.cell(r, first_open_col - 1).value or "").strip().upper()
            if not lbl and first_open_col >= 2:
                lbl = str(new_ws.cell(r, first_open_col - 2).value or "").strip().upper()

            has_dip_data = any(
                src_ws.cell(r, c_o).value is not None or src_ws.cell(r, c_c).value is not None
                for c_o, c_c in outer_dip_pairs
            )

            # Stop when encountering an empty row without label/data, or TOTAL / POS section
            if not has_dip_data and not lbl:
                break
            if any(k in lbl for k in ["TOTAL", "ADD UP", "ADDUP", "ZENITH", "OPAY", "PETRO"]):
                break

            row_had_update = False

            # Outer pairs reset
            for p_idx, (col_open, col_close) in enumerate(outer_dip_pairs):
                val_close = src_ws.cell(r, col_close).value
                val_open = src_ws.cell(r, col_open).value

                # If the cell contains an Excel formula (e.g. '=D20+D21' or '=E20+E21'), preserve it!
                if str(val_close or "").startswith("="):
                    continue

                if val_close is not None and (
                    isinstance(val_close, (int, float))
                    or (isinstance(val_close, str) and val_close.replace(".", "").isdigit())
                ):
                    last_dip = float(val_close) if isinstance(val_close, str) else val_close
                    new_ws.cell(r, col_open).value = last_dip
                    new_ws.cell(r, col_close).value = last_dip
                    row_had_update = True
                elif val_open is not None and (
                    isinstance(val_open, (int, float))
                    or (isinstance(val_open, str) and val_open.replace(".", "").isdigit())
                ):
                    last_dip = float(val_open) if isinstance(val_open, str) else val_open
                    new_ws.cell(r, col_open).value = last_dip
                    new_ws.cell(r, col_close).value = last_dip
                    row_had_update = True

            # Inner pairs reset
            for p_idx, (in_open, in_close) in enumerate(inner_dip_pairs):
                in_val_close = src_ws.cell(r, in_close).value
                in_val_open = src_ws.cell(r, in_open).value

                # If inner cell contains a formula (e.g. '=AC20+AC21'), preserve it!
                if str(in_val_close or "").startswith("="):
                    continue

                if in_val_close is not None and (
                    isinstance(in_val_close, (int, float))
                    or (isinstance(in_val_close, str) and in_val_close.replace(".", "").isdigit())
                ):
                    in_last_dip = float(in_val_close) if isinstance(in_val_close, str) else in_val_close
                    new_ws.cell(r, in_open).value = in_last_dip
                    new_ws.cell(r, in_close).value = in_last_dip
                    row_had_update = True
                elif in_val_open is not None and (
                    isinstance(in_val_open, (int, float))
                    or (isinstance(in_val_open, str) and in_val_open.replace(".", "").isdigit())
                ):
                    in_last_dip = float(in_val_open) if isinstance(in_val_open, str) else in_val_open
                    new_ws.cell(r, in_open).value = in_last_dip
                    new_ws.cell(r, in_close).value = in_last_dip
                    row_had_update = True
                elif p_idx < len(outer_dip_pairs):
                    # Fallback from outer if inner has no data
                    out_o, out_c = outer_dip_pairs[p_idx]
                    out_v = new_ws.cell(r, out_c).value
                    if out_v is not None and not str(out_v).startswith("="):
                        new_ws.cell(r, in_open).value = out_v
                        new_ws.cell(r, in_close).value = out_v

            if row_had_update:
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
            if val == "EXPENSES" or val.startswith("EXPENSES"):
                # Detect AMOUNT column for this expenses table (c + 1 or c + 2)
                amt_col = c + 1
                for c_amt in range(c + 1, min(c + 4, new_ws.max_column + 1)):
                    if "AMOUNT" in str(new_ws.cell(r, c_amt).value or "").upper():
                        amt_col = c_amt
                        break

                for r_sub in range(r + 1, min(r + 15, new_ws.max_row + 1)):
                    amt_cell = new_ws.cell(r_sub, amt_col)
                    lbl_cell = new_ws.cell(r_sub, c)
                    lbl_str = str(lbl_cell.value or "").strip().upper()
                    if "TOTAL" in lbl_str:
                        break
                    if amt_cell.value is not None and not str(amt_cell.value).startswith("="):
                        amt_cell.value = None
                        expenses_cleared += 1
                    if lbl_cell.value is not None and not str(lbl_cell.value).startswith("="):
                        lbl_cell.value = None
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
        # Detect ADD UP header on this row across full column width
        row_addup_cols = []
        row_avail_cols = []
        for c in range(1, new_ws.max_column + 1):
            val = str(new_ws.cell(r, c).value or "").strip().upper()
            val_norm = re.sub(r"[^A-Z0-9]", "", val)
            if "ADDUP" in val_norm or ("ADD" in val and "UP" in val):
                row_addup_cols.append(c)
            elif "AVAIL" in val:
                row_avail_cols.append(c)

        if row_addup_cols:
            header_row_1 = r
            val_row_1 = r + 1
            col_addup = row_addup_cols[0]

            # Find left col_avail_cash
            left_avails = [c for c in row_avail_cols if c < (row_addup_cols[1] if len(row_addup_cols) > 1 else 30)]
            col_avail_cash = left_avails[0] if left_avails else (col_addup + 1)

            # Left side POS columns (clear terminal values, skipping ADD UP and AVAILABLE CASH)
            left_boundary = max(col_addup, col_avail_cash)
            for c_pos in range(1, left_boundary + 1):
                if c_pos == col_addup or c_pos == col_avail_cash:
                    continue
                hdr_cell = new_ws.cell(header_row_1, c_pos).value
                if hdr_cell is not None and not str(hdr_cell).startswith("="):
                    cell_to_clear = new_ws.cell(val_row_1, c_pos)
                    if cell_to_clear.value is not None and not str(cell_to_clear.value).startswith("="):
                        cell_to_clear.value = None
                        pos_cleared += 1

            # Carry forward yesterday's (sheet -1) ending AVAILABLE CASH into ADD UP
            prev_avail_cash = _get_available_cash_value(src_ws, val_row_1, col_avail_cash, src_ws_data)

            # Fallback to existing ADD UP only if available cash could not be resolved at all
            if prev_avail_cash is None and col_addup <= src_ws.max_column:
                v_a = src_ws.cell(val_row_1, col_addup).value
                if isinstance(v_a, (int, float)):
                    prev_avail_cash = float(v_a)

            # Assign static float of previous day's AVAILABLE CASH to ADD UP (never an equation)
            if prev_avail_cash is not None:
                static_val = round(float(prev_avail_cash), 2)
                new_ws.cell(val_row_1, col_addup).value = static_val

                # Preserve formula on AVAILABLE CASH if it is an equation, else set static figure
                src_avail_cell = src_ws.cell(val_row_1, col_avail_cash).value
                if str(src_avail_cell or "").startswith("="):
                    new_ws.cell(val_row_1, col_avail_cash).value = src_avail_cell
                else:
                    new_ws.cell(val_row_1, col_avail_cash).value = static_val

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
                for c_pos in range(1, left_boundary + 1):
                    if c_pos == col_addup or c_pos == col_avail_cash:
                        continue
                    hdr_cell = new_ws.cell(header_row_2, c_pos).value
                    if hdr_cell is not None and not str(hdr_cell).startswith("="):
                        cell_to_clear = new_ws.cell(val_row_2, c_pos)
                        if cell_to_clear.value is not None and not str(cell_to_clear.value).startswith("="):
                            cell_to_clear.value = None
                            pos_cleared += 1

            # Right side POS columns (if inner POS table exists)
            right_addup_col = row_addup_cols[1] if len(row_addup_cols) > 1 else None
            right_avail_col = None
            if right_addup_col:
                right_avails = [c for c in row_avail_cols if c >= (right_addup_col - 5)]
                right_avail_col = right_avails[0] if right_avails else (right_addup_col + 1)

            min_right_c = left_boundary + 1
            end_right_pos = right_addup_col if right_addup_col else new_ws.max_column + 1
            for c_r in range(min_right_c, end_right_pos):
                if (right_addup_col and c_r == right_addup_col) or (right_avail_col and c_r == right_avail_col):
                    continue
                hdr_cell = new_ws.cell(header_row_1, c_r).value
                hdr_str = str(hdr_cell or "").strip().upper()
                if "ADD UP" in hdr_str or "AVAILABLE" in hdr_str:
                    continue
                if hdr_cell is not None and not str(hdr_cell).startswith("="):
                    cell_to_clear = new_ws.cell(val_row_1, c_r)
                    if cell_to_clear.value is not None and not str(cell_to_clear.value).startswith("="):
                        cell_to_clear.value = None
                        pos_cleared += 1

            # Ensure right-side ADD UP and AVAILABLE CASH figures are cleanly preserved (never an equation for ADD UP)
            for c_scan in range(left_boundary + 1, new_ws.max_column + 1):
                hdr_v = str(new_ws.cell(header_row_1, c_scan).value or "").strip().upper()
                if "ADD UP" in hdr_v:
                    src_val = src_ws.cell(val_row_1, c_scan).value
                    if isinstance(src_val, (int, float)):
                        new_ws.cell(val_row_1, c_scan).value = src_val
                    elif isinstance(src_val, str) and not src_val.startswith("="):
                        try:
                            new_ws.cell(val_row_1, c_scan).value = float(src_val.replace(",", "").strip())
                        except ValueError:
                            new_ws.cell(val_row_1, c_scan).value = src_val
                elif "AVAILABLE" in hdr_v:
                    src_avail = src_ws.cell(val_row_1, c_scan).value
                    if str(src_avail or "").startswith("="):
                        new_ws.cell(val_row_1, c_scan).value = src_avail
                    elif isinstance(src_avail, (int, float)):
                        new_ws.cell(val_row_1, c_scan).value = src_avail
                    elif isinstance(src_avail, str) and not src_avail.startswith("="):
                        try:
                            new_ws.cell(val_row_1, c_scan).value = float(src_avail.replace(",", "").strip())
                        except ValueError:
                            new_ws.cell(val_row_1, c_scan).value = src_avail

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

    try:
        wb.calculation.fullCalcOnLoad = True
    except Exception:
        pass

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
        try:
            wb_data = openpyxl.load_workbook(spooled_path, data_only=True)
        except Exception:
            wb_data = None
        updated_wb, summary = sync_station_report(wb, wb_data=wb_data)

        # Infer clean station name from filename
        clean_stem = clean_report_name(orig_name)
        summary.filename = orig_name
        summary.station_name = clean_stem

        buf = io.BytesIO()
        updated_wb.save(buf)
        buf.seek(0)
        output_bytes = buf.getvalue()

        # Output filename contains "...sync..."
        out_name = f"{clean_stem}_sync.xlsx"
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
