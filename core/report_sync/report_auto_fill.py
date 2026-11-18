"""Core engine for Report Auto-Fill (Outer to Inner Report Sync).

After a station daily report has been filled on the Outer Meter (Left side),
this engine copies and synchronizes the corresponding operational figures to the Inner Meter (Right side):
1. Meters: Heals zero-meter corruption via backward historical lookback and aligns inner closing meters.
2. Tank Dipping: Symmetrical multi-pair dynamic dipping copy (Outer -> Inner), strictly preserving Excel formulas.
3. POS List: Terminal matching across Day N + Day N+1 rollover rows, maintaining static ADD UP figures.
4. Products Received: Outer -> Inner litres received copy up to row 25.
5. Expenses & Imprest: Outer -> Inner description, amount, and remittance imprest copy.
6. Single-sided station detection (max_column < 30) reporting clean success.
7. Production clean filename pipeline: stripping all '_sync', '_synced', '_filled', timestamps.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import io
import logging
from pathlib import Path
import re
from typing import Any, Callable
import zipfile

import openpyxl

from core.memory import cleanup_files, force_gc, spool_uploaded_file
from .report_syncer import _resolve_last_active_closing, clean_report_name

logger = logging.getLogger(__name__)


@dataclass
class AutoFillSummary:
    filename: str
    station_name: str
    sheet_name: str
    date_str: str
    tanks_copied: int
    pos_copied: int
    products_copied: int
    expenses_copied: int
    pumps_synced: int = 0
    status: str = "Success"
    message: str = "Report successfully filled from Left side to Right side."


def _norm_label(s: Any) -> str:
    """Normalizes a label string for robust matching across stations."""
    return re.sub(r"[^A-Z0-9]", "", str(s or "").upper())


def auto_fill_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    wb_worksheets: list[openpyxl.worksheet.worksheet.Worksheet] | None = None,
) -> dict[str, int]:
    """Copies Tank Dipping, POS list, Products Received, Expenses, and synchronizes pump meters from Left to Right.

    Args:
        ws: openpyxl worksheet to process.
        wb_worksheets: Optional list of all worksheets in workbook for historical meter recovery.

    Returns:
        dict with counts of items copied/healed: tanks, pos, products, expenses, pumps.
    """
    counts = {
        "tanks": 0,
        "pos": 0,
        "products": 0,
        "expenses": 0,
        "pumps": 0,
    }

    if wb_worksheets is None:
        wb_worksheets = getattr(ws, "parent", None) and ws.parent.worksheets or [ws]

    # Handle single-sided stations (e.g. Jah michael with no inner meter table)
    if ws.max_column < 30:
        return counts

    # ---------------------------------------------------------
    # 1. TANK TABLE HEADER & DIP PAIRS (Multi-pair dynamic tank support)
    # ---------------------------------------------------------
    tank_header_row = None
    dip_pairs: list[tuple[int, int]] = []

    for r in range(8, min(45, ws.max_row + 1)):
        row_pairs = []
        for c in range(1, ws.max_column):
            v1 = str(ws.cell(r, c).value or "").strip().upper()
            v2 = str(ws.cell(r, c + 1).value or "").strip().upper()
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
        for r in range(8, min(45, ws.max_row + 1)):
            has_tank = any("TANK" in str(ws.cell(r, c).value or "").upper() for c in range(1, ws.max_column + 1))
            if has_tank:
                row_pairs = []
                for c in range(1, ws.max_column):
                    v1 = str(ws.cell(r, c).value or "").strip().upper()
                    v2 = str(ws.cell(r, c + 1).value or "").strip().upper()
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
        outer_dip_pairs = [p for p in dip_pairs if p[0] - first_c < 10]
        inner_dip_pairs = [p for p in dip_pairs if p[0] - first_c >= 10]

    # ---------------------------------------------------------
    # 2. METER TABLE HEADERS & PUMP HEALING / ALIGNMENT
    # ---------------------------------------------------------
    meter_header_row = None
    outer_open_col = None
    inner_open_col = None
    outer_close_cols: list[int] = []
    inner_close_cols: list[int] = []

    for r in range(1, 6):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(r, c).value or "").strip().upper()
            if "OPENING METER" in val or "OPEN METER" in val:
                if outer_open_col is None:
                    outer_open_col = c
                    meter_header_row = r
                elif c > outer_open_col + 5:
                    inner_open_col = c

    # Fallback meter header detection if "OPENING METER" keyword is mislabeled
    if outer_open_col is None:
        for r in range(1, 6):
            meter_cols = []
            for c in range(1, ws.max_column + 1):
                v = str(ws.cell(r, c).value or "").strip().upper()
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
            for c in range(1, min(25, ws.max_column + 1)):
                val = str(ws.cell(r, c).value or "").strip().upper()
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

    if inner_open_col is None and meter_header_row and ws.max_column > 25:
        for r in (3, 4):
            if inner_open_col:
                break
            for c in range(25, ws.max_column + 1):
                val = str(ws.cell(r, c).value or "").strip().upper()
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
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(meter_header_row, c).value or "").strip().upper()
            if "CLOSING METER" in val or "CLOSE METER" in val:
                if inner_open_col and c > inner_open_col:
                    inner_close_cols.append(c)
                else:
                    outer_close_cols.append(c)

    meter_scan_limit = tank_header_row if tank_header_row else min(25, ws.max_row + 1)

    if meter_header_row and outer_open_col:
        for r in range(meter_header_row + 1, meter_scan_limit):
            row_vals_sample = [str(ws.cell(r, c).value or "").strip().upper() for c in range(1, min(10, ws.max_column + 1))]
            if any(v == "TANK" or v.startswith("TANK") for v in row_vals_sample):
                break

            out_lbl = str(ws.cell(r, outer_open_col - 1).value or "").strip().upper()
            if "TOTAL" in out_lbl or out_lbl.startswith("="):
                continue

            val_out_open = ws.cell(r, outer_open_col).value
            is_num_out = isinstance(val_out_open, (int, float)) or (
                isinstance(val_out_open, str) and val_out_open.replace(".", "").isdigit()
            )

            # Heal outer meter if 0 or None via lookback
            if (not is_num_out or val_out_open == 0 or val_out_open == "0") and wb_worksheets:
                last_cl_rec, any_p = _resolve_last_active_closing(
                    wb_worksheets, r, outer_open_col, outer_close_cols
                )
                if any_p or (last_cl_rec is not None and last_cl_rec > 0):
                    ws.cell(r, outer_open_col).value = last_cl_rec
                    val_out_open = last_cl_rec
                    is_num_out = True
                    counts["pumps"] += 1

            # Determine outer sales
            outer_sales = 0.0
            last_out_close = None
            if outer_close_cols:
                for c in reversed(outer_close_cols):
                    v_c = ws.cell(r, c).value
                    if v_c is not None and isinstance(v_c, (int, float)) and v_c > 0 and not str(v_c).startswith("="):
                        last_out_close = v_c
                        break
                    if isinstance(v_c, str) and v_c.replace(".", "").isdigit() and float(v_c) > 0:
                        last_out_close = float(v_c) if "." in v_c else int(v_c)
                        break

            if is_num_out and last_out_close is not None and isinstance(val_out_open, (int, float)):
                if last_out_close > val_out_open:
                    outer_sales = float(last_out_close) - float(val_out_open)

            # Inner meter healing and auto-fill
            if inner_open_col:
                in_lbl = str(ws.cell(r, inner_open_col - 1).value or "").strip().upper()
                if "TOTAL" in in_lbl or in_lbl.startswith("="):
                    continue

                val_in_open = ws.cell(r, inner_open_col).value
                is_num_in = isinstance(val_in_open, (int, float)) or (
                    isinstance(val_in_open, str) and val_in_open.replace(".", "").isdigit()
                )

                if (not is_num_in or val_in_open == 0 or val_in_open == "0") and wb_worksheets:
                    last_in_rec, in_any_p = _resolve_last_active_closing(
                        wb_worksheets, r, inner_open_col, inner_close_cols
                    )
                    if in_any_p or (last_in_rec is not None and last_in_rec > 0):
                        ws.cell(r, inner_open_col).value = last_in_rec
                        val_in_open = last_in_rec
                        is_num_in = True
                        counts["pumps"] += 1

                # Auto-fill inner closing meter if not yet filled by auditor
                if is_num_in and inner_close_cols and outer_sales > 0:
                    target_in_col = inner_close_cols[0]
                    val_in_close = ws.cell(r, target_in_col).value
                    in_needs_fill = (
                        val_in_close is None
                        or val_in_close == 0
                        or val_in_close == "0"
                        or val_in_close == val_in_open
                    )
                    if in_needs_fill:
                        if (
                            isinstance(val_in_open, (int, float))
                            and isinstance(val_out_open, (int, float))
                            and abs(float(val_in_open) - float(val_out_open)) < 1
                            and last_out_close is not None
                        ):
                            ws.cell(r, target_in_col).value = last_out_close
                        else:
                            new_in_close = round(float(val_in_open) + outer_sales)
                            ws.cell(r, target_in_col).value = new_in_close
                        counts["pumps"] += 1

    # ---------------------------------------------------------
    # 3. TANK DIPPING: Outer -> Inner (Multi-pair dynamic tank support)
    # ---------------------------------------------------------
    if tank_header_row and outer_dip_pairs and inner_dip_pairs:
        first_open_col = outer_dip_pairs[0][0]
        for r in range(tank_header_row + 1, min(tank_header_row + 15, ws.max_row + 1)):
            lbl = ws.cell(r, first_open_col - 1).value
            if not lbl and first_open_col >= 2:
                lbl = ws.cell(r, first_open_col - 2).value
            lbl_str = str(lbl or "").strip().upper()

            has_dip_data = any(
                ws.cell(r, c_o).value is not None or ws.cell(r, c_c).value is not None
                for c_o, c_c in outer_dip_pairs
            )

            # Boundary checks: Stop if empty, TOTAL, or POS table
            if not has_dip_data and not lbl:
                break
            if any(k in lbl_str for k in ["TOTAL", "ADD UP", "ADDUP", "ZENITH", "OPAY", "PETRO"]):
                break

            row_copied = False
            for p_idx in range(min(len(outer_dip_pairs), len(inner_dip_pairs))):
                c_o, c_c = outer_dip_pairs[p_idx]
                in_o, in_c = inner_dip_pairs[p_idx]

                v_open = ws.cell(r, c_o).value
                v_close = ws.cell(r, c_c).value
                in_v_open = ws.cell(r, in_o).value
                in_v_close = ws.cell(r, in_c).value

                # Preserve Excel formulas on both inner and outer tables
                in_o_formula = str(in_v_open or "").startswith("=")
                in_c_formula = str(in_v_close or "").startswith("=")
                out_o_formula = str(v_open or "").startswith("=")
                out_c_formula = str(v_close or "").startswith("=")

                if not in_o_formula and not out_o_formula and (v_open is not None and v_open != ""):
                    ws.cell(r, in_o).value = v_open
                    row_copied = True

                if not in_c_formula and not out_c_formula and (v_close is not None and v_close != ""):
                    ws.cell(r, in_c).value = v_close
                    row_copied = True

            if row_copied:
                counts["tanks"] += 1

    # ---------------------------------------------------------
    # 4. PRODUCTS RECEIVED: Outer -> Inner
    # ---------------------------------------------------------
    outer_prod_col = None
    inner_prod_col = None
    prod_header_row_outer = None
    prod_header_row_inner = None

    for r in range(1, min(25, ws.max_row + 1)):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(r, c).value or "").strip().upper()
            if "PRODUCTS RECEIVED" in val or "PRODUCT RECEIVED" in val:
                if c < 30 and outer_prod_col is None:
                    outer_prod_col = c
                    prod_header_row_outer = r
                elif c >= 30 and inner_prod_col is None:
                    inner_prod_col = c
                    prod_header_row_inner = r

    if outer_prod_col and inner_prod_col:
        p_out = prod_header_row_outer or 6
        p_in = prod_header_row_inner or p_out
        for offset in range(1, 12):
            r_out = p_out + offset
            r_in = p_in + offset
            if r_out > ws.max_row or r_in > ws.max_row:
                break
            lbl_outer = ws.cell(r_out, outer_prod_col).value
            lbl_str = str(lbl_outer or "").strip().upper()

            if "TOTAL" in lbl_str or "EXPENSES" in lbl_str:
                break
            if not lbl_outer and ws.cell(r_out, outer_prod_col + 1).value is None and ws.cell(r_in, inner_prod_col + 1).value is None:
                break

            litres_val = ws.cell(r_out, outer_prod_col + 1).value
            if litres_val is not None and not str(litres_val).startswith("="):
                ws.cell(r_in, inner_prod_col + 1).value = litres_val
                counts["products"] += 1

    # ---------------------------------------------------------
    # 5. EXPENSES & IMPREST: Outer -> Inner
    # ---------------------------------------------------------
    outer_exp_col = None
    inner_exp_col = None
    exp_header_row_outer = None
    exp_header_row_inner = None

    for r in range(5, min(28, ws.max_row + 1)):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(r, c).value or "").strip().upper()
            if val == "EXPENSES" or ("EXPENSES" in val and not str(val).startswith("=")):
                if c < 30 and outer_exp_col is None:
                    outer_exp_col = c
                    exp_header_row_outer = r
                elif c >= 30 and inner_exp_col is None:
                    inner_exp_col = c
                    exp_header_row_inner = r

    if outer_exp_col and inner_exp_col:
        e_out = exp_header_row_outer or 12
        e_in = exp_header_row_inner or e_out
        for offset in range(1, 15):
            r_out = e_out + offset
            r_in = e_in + offset
            if r_out > ws.max_row or r_in > ws.max_row:
                break
            lbl_outer = ws.cell(r_out, outer_exp_col).value
            amt_outer = ws.cell(r_out, outer_exp_col + 1).value
            lbl_str = str(lbl_outer or "").strip().upper()

            if "TOTAL" in lbl_str:
                break
            if not lbl_outer and amt_outer is None:
                break

            if not str(amt_outer or "").startswith("="):
                ws.cell(r_in, inner_exp_col).value = lbl_outer
                ws.cell(r_in, inner_exp_col + 1).value = amt_outer
                counts["expenses"] += 1

    # Imprest in Remittance Report (Outer -> Inner)
    outer_imprest_val = None
    inner_imprest_cell = None
    for r in range(3, min(15, ws.max_row + 1)):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(r, c).value or "").strip().upper()
            if val == "IMPREST":
                if c < 30:
                    for c_scan in (c + 1, c + 2):
                        v_sc = ws.cell(r, c_scan).value
                        if v_sc is not None and not str(v_sc).startswith("="):
                            outer_imprest_val = v_sc
                            break
                else:
                    for c_scan in (c + 1, c + 2):
                        inner_imprest_cell = ws.cell(r, c_scan)
                        break

    if outer_imprest_val is not None and inner_imprest_cell is not None:
        inner_imprest_cell.value = outer_imprest_val
        counts["expenses"] += 1

    # ---------------------------------------------------------
    # 6. POS LIST: Matching Left-Side Terminals -> Right-Side Terminals
    # ---------------------------------------------------------
    for r in range(15, ws.max_row + 1):
        for c in range(1, min(30, ws.max_column + 1)):
            val = str(ws.cell(r, c).value or "").strip().upper()
            val_norm = re.sub(r"[^A-Z0-9]", "", val)
            if "ADDUP" in val_norm or ("ADD" in val and "UP" in val):
                header_row_1 = r
                val_row_1 = r + 1
                col_addup = c

                col_avail_cash = c + 1
                for c_scan in range(c + 1, min(c + 5, ws.max_column + 1)):
                    if "AVAILABLE" in str(ws.cell(header_row_1, c_scan).value or "").upper():
                        col_avail_cash = c_scan
                        break

                val_row_2 = header_row_1 + 3 if header_row_1 + 3 <= ws.max_row else None

                # Build dictionary of Left-Side POS terminal values
                left_pos: dict[str, Any] = {}
                for c_l in range(1, col_addup):
                    h_l = ws.cell(header_row_1, c_l).value
                    if h_l and not str(h_l).startswith("="):
                        k = _norm_label(h_l)
                        v1 = ws.cell(val_row_1, c_l).value
                        v2 = ws.cell(val_row_2, c_l).value if val_row_2 else None

                        if isinstance(v1, (int, float)) and isinstance(v2, (int, float)):
                            left_pos[k] = v1 + v2
                        elif isinstance(v1, (int, float)):
                            left_pos[k] = v1
                        elif isinstance(v2, (int, float)):
                            left_pos[k] = v2
                        else:
                            left_pos[k] = v1

                # Find right-side ADD UP boundary
                right_addup_col = None
                for c_r in range(col_avail_cash + 1, ws.max_column + 1):
                    v_r = str(ws.cell(header_row_1, c_r).value or "").strip().upper()
                    v_r_norm = re.sub(r"[^A-Z0-9]", "", v_r)
                    if "ADDUP" in v_r_norm or ("ADD" in v_r and "UP" in v_r):
                        right_addup_col = c_r
                        break

                end_pos_col = right_addup_col if right_addup_col else ws.max_column + 1

                # Copy matching terminals to Right Side
                for c_r in range(col_avail_cash + 1, end_pos_col):
                    h_r = ws.cell(header_row_1, c_r).value
                    h_r_str = str(h_r or "").strip().upper()
                    if not h_r or "ADD" in h_r_str or "AVAILABLE" in h_r_str or str(h_r).startswith("="):
                        continue

                    k_r = _norm_label(h_r)
                    if k_r in left_pos:
                        match_val = left_pos[k_r]
                        ws.cell(val_row_1, c_r).value = match_val
                        counts["pos"] += 1

                # Maintain exact static figure on right ADD UP (never an equation)
                left_addup_val = ws.cell(val_row_1, col_addup).value
                if right_addup_col and left_addup_val is not None:
                    if isinstance(left_addup_val, (int, float)):
                        ws.cell(val_row_1, right_addup_col).value = left_addup_val
                    elif isinstance(left_addup_val, str) and not left_addup_val.startswith("="):
                        try:
                            ws.cell(val_row_1, right_addup_col).value = float(left_addup_val.replace(",", "").strip())
                        except ValueError:
                            pass

                # Stop after processing the POS table
                break

    return counts


def auto_fill_workbook(
    wb: openpyxl.Workbook,
    target_sheet_mode: str = "latest",
    specific_sheet_name: str | None = None,
) -> tuple[openpyxl.Workbook, list[AutoFillSummary]]:
    """Runs auto-fill across specified sheet(s) in a workbook."""
    if not wb.worksheets:
        raise ValueError("Workbook contains no worksheets.")

    sheets_to_process = []
    if target_sheet_mode == "latest":
        sheets_to_process = [wb.worksheets[-1]]
    elif target_sheet_mode == "specific" and specific_sheet_name:
        if specific_sheet_name in wb.sheetnames:
            sheets_to_process = [wb[specific_sheet_name]]
        else:
            raise ValueError(f"Sheet '{specific_sheet_name}' not found in workbook.")
    else:  # "all"
        sheets_to_process = wb.worksheets

    summaries = []
    for ws in sheets_to_process:
        # Extract date (scan rows 1 to 3)
        date_str = "Unknown"
        for r_d in range(1, 4):
            for c in range(1, min(10, ws.max_column + 1)):
                v = ws.cell(r_d, c).value
                if isinstance(v, datetime):
                    date_str = v.strftime("%Y-%m-%d")
                    break
                elif v is not None and any(char.isdigit() for char in str(v)):
                    date_str = str(v)[:10]
                    break
            if date_str != "Unknown":
                break

        counts = auto_fill_sheet(ws, wb_worksheets=wb.worksheets)
        is_single_sided = ws.max_column < 30
        if is_single_sided:
            msg = "Single-sided station report (no inner meter table). Preserved."
        else:
            msg = f"Copied {counts['tanks']} tanks, {counts['pos']} POS terminals, {counts['products']} products, {counts['expenses']} expenses."
            if counts.get("pumps", 0) > 0:
                msg += f" Synced/healed {counts['pumps']} pump meters."

        summaries.append(
            AutoFillSummary(
                filename="",
                station_name="",
                sheet_name=ws.title,
                date_str=date_str,
                tanks_copied=counts["tanks"],
                pos_copied=counts["pos"],
                products_copied=counts["products"],
                expenses_copied=counts["expenses"],
                pumps_synced=counts.get("pumps", 0),
                status="Success",
                message=msg,
            )
        )

    try:
        wb.calculation.fullCalcOnLoad = True
    except Exception:
        pass

    return wb, summaries


def auto_fill_station_file(
    input_file: Any,
    target_sheet_mode: str = "latest",
    specific_sheet_name: str | None = None,
) -> tuple[bytes, str, list[AutoFillSummary]]:
    """Processes a single station file and returns updated bytes, output filename, and summaries."""
    if isinstance(input_file, (str, Path)):
        orig_name = Path(input_file).name
    else:
        orig_name = getattr(input_file, "name", "station_report.xlsx")

    spooled_path = spool_uploaded_file(input_file, prefix="auto_fill_")

    try:
        wb = openpyxl.load_workbook(spooled_path, data_only=False)
        updated_wb, summaries = auto_fill_workbook(
            wb,
            target_sheet_mode=target_sheet_mode,
            specific_sheet_name=specific_sheet_name,
        )

        clean_stem = clean_report_name(orig_name)
        for s in summaries:
            s.filename = orig_name
            s.station_name = clean_stem

        buf = io.BytesIO()
        updated_wb.save(buf)
        buf.seek(0)
        output_bytes = buf.getvalue()

        # User requirement: output must come out clean as the report name
        out_name = f"{clean_stem}.xlsx"
        return output_bytes, out_name, summaries

    finally:
        cleanup_files(spooled_path)
        force_gc()


def auto_fill_multiple_station_files(
    files: list[Any],
    target_sheet_mode: str = "latest",
    specific_sheet_name: str | None = None,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[list[tuple[str, bytes]], io.BytesIO, list[AutoFillSummary]]:
    """Batch processes multiple station workbooks for auto-fill and packages into a ZIP archive."""
    results: list[tuple[str, bytes]] = []
    all_summaries: list[AutoFillSummary] = []
    total = len(files)

    for idx, f in enumerate(files):
        name = Path(f).name if isinstance(f, (str, Path)) else getattr(f, "name", f"file_{idx+1}.xlsx")
        if progress_callback:
            progress_callback(idx, total, name)

        try:
            out_bytes, out_name, summaries = auto_fill_station_file(
                f,
                target_sheet_mode=target_sheet_mode,
                specific_sheet_name=specific_sheet_name,
            )
            results.append((out_name, out_bytes))
            all_summaries.extend(summaries)
        except Exception as e:
            logger.error("Failed to auto-fill %s: %s", name, e, exc_info=True)
            all_summaries.append(
                AutoFillSummary(
                    filename=name,
                    station_name=Path(name).stem,
                    sheet_name="Error",
                    date_str="Error",
                    tanks_copied=0,
                    pos_copied=0,
                    products_copied=0,
                    expenses_copied=0,
                    pumps_synced=0,
                    status="Error",
                    message=str(e),
                )
            )

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
    return results, zip_buf, all_summaries
