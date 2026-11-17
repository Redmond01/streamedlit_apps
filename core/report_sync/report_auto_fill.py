"""Core engine for Report Auto-Fill (Outer to Inner Report Sync).

After a station daily report has been filled on the Outer Meter (Left side),
this engine copies the corresponding operational figures to the Inner Meter (Right side):
1. Tank Dipping (Opening Dipping, Closing Dipping)
2. POS List (Matching terminal collections from Left side to Right side)
3. Products Received (Litres received)
4. Expenses (Descriptions and amounts, including imprest)
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
    status: str = "Success"
    message: str = "Report successfully filled from Left side to Right side."


def _norm_label(s: Any) -> str:
    """Normalizes a label string for robust matching across stations."""
    return re.sub(r"[^A-Z0-9]", "", str(s or "").upper())


def auto_fill_sheet(ws: openpyxl.worksheet.worksheet.Worksheet) -> dict[str, int]:
    """Copies Tank Dipping, POS list, Products Received, and Expenses from Left to Right side on a worksheet.

    Returns:
        dict with counts of items copied: tanks, pos, products, expenses.
    """
    counts = {
        "tanks": 0,
        "pos": 0,
        "products": 0,
        "expenses": 0,
    }

    # ---------------------------------------------------------
    # 1. TANK DIPPING: Outer -> Inner
    # ---------------------------------------------------------
    tank_header_row = None
    outer_tank_open_col = None
    outer_tank_close_col = None
    inner_tank_open_col = None
    inner_tank_close_col = None

    for r in range(10, min(30, ws.max_row + 1)):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(r, c).value or "").strip().upper()
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

    if tank_header_row and outer_tank_open_col and outer_tank_close_col and inner_tank_open_col and inner_tank_close_col:
        for r in range(tank_header_row + 1, min(tank_header_row + 15, ws.max_row + 1)):
            lbl = ws.cell(r, outer_tank_open_col - 1).value
            lbl_str = str(lbl or "").strip().upper()

            # Boundary checks: Stop if empty, TOTAL, or POS table
            if not lbl and ws.cell(r, outer_tank_open_col).value is None:
                break
            if "TOTAL" in lbl_str or "ADD UP" in lbl_str or "ZENITH" in lbl_str or "OPAY" in lbl_str or "PETRO" in lbl_str:
                break

            v_open = ws.cell(r, outer_tank_open_col).value
            v_close = ws.cell(r, outer_tank_close_col).value

            # If opening or closing dipping is present, copy to inner dipping table
            if v_open is not None or v_close is not None:
                ws.cell(r, inner_tank_open_col).value = v_open
                ws.cell(r, inner_tank_close_col).value = v_close
                counts["tanks"] += 1

    # ---------------------------------------------------------
    # 2. PRODUCTS RECEIVED: Outer -> Inner
    # ---------------------------------------------------------
    outer_prod_col = None
    inner_prod_col = None
    prod_header_row = None

    for r in range(1, min(25, ws.max_row + 1)):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(r, c).value or "").strip().upper()
            if "PRODUCTS RECEIVED" in val or "PRODUCT RECEIVED" in val:
                if outer_prod_col is None:
                    outer_prod_col = c
                    prod_header_row = r
                else:
                    inner_prod_col = c

    if prod_header_row and outer_prod_col and inner_prod_col:
        for r in range(prod_header_row + 1, min(prod_header_row + 12, ws.max_row + 1)):
            lbl_outer = ws.cell(r, outer_prod_col).value
            lbl_str = str(lbl_outer or "").strip().upper()

            if "TOTAL" in lbl_str or "EXPENSES" in lbl_str:
                break
            if not lbl_outer and ws.cell(r, outer_prod_col + 1).value is None:
                break

            litres_val = ws.cell(r, outer_prod_col + 1).value
            if litres_val is not None:
                ws.cell(r, inner_prod_col + 1).value = litres_val
                counts["products"] += 1

    # ---------------------------------------------------------
    # 3. EXPENSES & IMPREST: Outer -> Inner
    # ---------------------------------------------------------
    outer_exp_col = None
    inner_exp_col = None
    exp_header_row = None

    for r in range(5, min(28, ws.max_row + 1)):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(r, c).value or "").strip().upper()
            if val == "EXPENSES":
                if outer_exp_col is None:
                    outer_exp_col = c
                    exp_header_row = r
                else:
                    inner_exp_col = c

    if exp_header_row and outer_exp_col and inner_exp_col:
        for r in range(exp_header_row + 1, min(exp_header_row + 15, ws.max_row + 1)):
            lbl_outer = ws.cell(r, outer_exp_col).value
            amt_outer = ws.cell(r, outer_exp_col + 1).value
            lbl_str = str(lbl_outer or "").strip().upper()

            if "TOTAL" in lbl_str:
                break
            if not lbl_outer and amt_outer is None:
                break

            if lbl_outer is not None or amt_outer is not None:
                ws.cell(r, inner_exp_col).value = lbl_outer
                ws.cell(r, inner_exp_col + 1).value = amt_outer
                counts["expenses"] += 1

    # Check Imprest in Remittance Report (Outer -> Inner)
    outer_imprest_val = None
    inner_imprest_cell = None
    for r in range(3, min(15, ws.max_row + 1)):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(r, c).value or "").strip().upper()
            if val == "IMPREST":
                if c < 30:
                    # Outer imprest
                    for c_scan in (c + 1, c + 2):
                        v_sc = ws.cell(r, c_scan).value
                        if v_sc is not None and not str(v_sc).startswith("="):
                            outer_imprest_val = v_sc
                            break
                else:
                    # Inner imprest
                    for c_scan in (c + 1, c + 2):
                        inner_imprest_cell = ws.cell(r, c_scan)
                        break

    if outer_imprest_val is not None and inner_imprest_cell is not None:
        inner_imprest_cell.value = outer_imprest_val
        counts["expenses"] += 1

    # ---------------------------------------------------------
    # 4. POS LIST: Matching Left-Side Terminals -> Right-Side Terminals
    # ---------------------------------------------------------
    for r in range(15, ws.max_row + 1):
        for c in range(1, min(30, ws.max_column + 1)):
            val = str(ws.cell(r, c).value or "").strip().upper()
            if "ADD UP" in val:
                header_row_1 = r
                val_row_1 = r + 1
                col_addup = c
                col_avail_cash = c + 1

                # Check for rollover row
                val_row_2 = header_row_1 + 3 if header_row_1 + 3 <= ws.max_row else None

                # Build dictionary of Left-Side POS terminal values
                left_pos: dict[str, Any] = {}
                for c_l in range(1, col_addup):
                    h_l = ws.cell(header_row_1, c_l).value
                    if h_l and not str(h_l).startswith("="):
                        k = _norm_label(h_l)
                        v1 = ws.cell(val_row_1, c_l).value
                        v2 = ws.cell(val_row_2, c_l).value if val_row_2 else None

                        # Combine Day N + Rollover Day N+1 if numeric
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
                    if "ADD UP" in v_r:
                        right_addup_col = c_r
                        break

                end_pos_col = right_addup_col if right_addup_col else ws.max_column + 1

                # Copy matching terminals to Right Side
                for c_r in range(col_avail_cash + 1, end_pos_col):
                    h_r = ws.cell(header_row_1, c_r).value
                    h_r_str = str(h_r or "").strip().upper()
                    if not h_r or "ADD UP" in h_r_str or "AVAILABLE" in h_r_str or str(h_r).startswith("="):
                        continue

                    k_r = _norm_label(h_r)
                    if k_r in left_pos:
                        match_val = left_pos[k_r]
                        ws.cell(val_row_1, c_r).value = match_val
                        counts["pos"] += 1

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
        # Extract date
        date_str = "Unknown"
        for c in range(1, min(10, ws.max_column + 1)):
            v = ws.cell(1, c).value
            if isinstance(v, datetime):
                date_str = v.strftime("%Y-%m-%d")
                break
            elif v is not None and any(char.isdigit() for char in str(v)):
                date_str = str(v)[:10]
                break

        counts = auto_fill_sheet(ws)
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
                status="Success",
                message=f"Copied {counts['tanks']} tanks, {counts['pos']} POS terminals, {counts['products']} products, {counts['expenses']} expenses.",
            )
        )

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

        stem = Path(orig_name).stem
        for s in summaries:
            s.filename = orig_name
            s.station_name = stem

        buf = io.BytesIO()
        updated_wb.save(buf)
        buf.seek(0)
        output_bytes = buf.getvalue()

        out_name = f"{stem}_filled_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
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
