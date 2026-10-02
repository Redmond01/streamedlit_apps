from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
import logging
import re
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook

from core.gear.utils import is_blank, parse_number
from .config import SalesConfig

logger = logging.getLogger(__name__)

DATE_REGEX = re.compile(r"(\d{4}[-/]\d{1,2}[-/]\d{1,2})|(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})")


@dataclass(frozen=True)
class DailySalesRecord:
    station_file: str
    sheet_name: str
    date: date
    pms: Decimal | None
    ago: Decimal | None
    lpg: Decimal | None


MONTH_NAME_TO_INT: dict[str, int] = {
    "JANUARY": 1, "JAN": 1,
    "FEBRUARY": 2, "FEB": 2,
    "MARCH": 3, "MAR": 3,
    "APRIL": 4, "APR": 4,
    "MAY": 5,
    "JUNE": 6, "JUN": 6,
    "JULY": 7, "JUL": 7,
    "AUGUST": 8, "AUG": 8,
    "SEPTEMBER": 9, "SEP": 9, "SEPT": 9,
    "OCTOBER": 10, "OCT": 10,
    "NOVEMBER": 11, "NOV": 11,
    "DECEMBER": 12, "DEC": 12,
}


def resolve_inverted_date(
    d: date,
    filename: str | None = None,
    target_date: date | None = None,
    expected_month: int | None = None,
) -> date:
    """Detects and corrects dates where Day and Month were inverted by Excel.

    For example, typing '01/10/2026' into a sheet with mm-dd-yy format causes Excel
    to store it as 2026-01-10 (Jan 10) instead of 2026-10-01 (Oct 1).
    """
    if not isinstance(d, date):
        return d

    # 1. Exact match with target_date when day & month are swapped
    if target_date and d.year == target_date.year:
        if d.day == target_date.month and d.month == target_date.day:
            return target_date

    # 2. Determine target month from expected_month or filename
    inferred_month = expected_month
    if inferred_month is None and target_date:
        inferred_month = target_date.month
    if inferred_month is None and filename:
        fn_upper = filename.upper()
        for mname, mnum in MONTH_NAME_TO_INT.items():
            if re.search(r"\b" + mname + r"\b", fn_upper):
                inferred_month = mnum
                break

    # 3. If date's month does not match inferred month, but day does, swap them
    if inferred_month and d.month != inferred_month:
        if d.day == inferred_month and 1 <= d.month <= 12:
            try:
                corrected = date(d.year, d.day, d.month)
                logger.info(
                    "Corrected inverted date from %s to %s (inferred month: %d, file: %s)",
                    d, corrected, inferred_month, filename
                )
                return corrected
            except ValueError:
                pass

    return d


def extract_date_from_sheet(
    sheet: Worksheet,
    filename: str | None = None,
    target_date: date | None = None,
    expected_month: int | None = None,
) -> date | None:
    """Extracts calendar date from daily station worksheet, resolving inverted dates."""
    max_r = min(sheet.max_row or 10, 8)
    max_c = min(sheet.max_column or 15, 12)

    raw_date: date | None = None

    # Pass 1: Direct datetime/date objects in top rows
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            val = sheet.cell(r, c).value
            if isinstance(val, (datetime, date)):
                raw_date = val.date() if isinstance(val, datetime) else val
                break
        if raw_date:
            break

    # Pass 2: Look for 'DATE' label
    if not raw_date:
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                val = str(sheet.cell(r, c).value or "").strip().upper()
                if val in {"DATE", "DATE:", "TRANS DATE", "TRANS DATE:"}:
                    for check_cell in [
                        sheet.cell(r, c + 1).value,
                        sheet.cell(r + 1, c).value,
                        sheet.cell(r, c + 2).value,
                    ]:
                        if isinstance(check_cell, (datetime, date)):
                            raw_date = check_cell.date() if isinstance(check_cell, datetime) else check_cell
                            break
                        if isinstance(check_cell, str):
                            parsed = _parse_date_str(check_cell)
                            if parsed:
                                raw_date = parsed
                                break
                    if raw_date:
                        break
            if raw_date:
                break

    # Pass 3: Regex string matching across top cells
    if not raw_date:
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                val = sheet.cell(r, c).value
                if isinstance(val, str):
                    parsed = _parse_date_str(val)
                    if parsed:
                        raw_date = parsed
                        break
            if raw_date:
                break

    # Pass 4: Fallback to sheet title
    if not raw_date:
        raw_date = _parse_date_str(sheet.title)

    if raw_date:
        return resolve_inverted_date(
            raw_date,
            filename=filename,
            target_date=target_date,
            expected_month=expected_month,
        )

    return None


def _parse_date_str(text: str) -> date | None:
    if not text:
        return None
    match = DATE_REGEX.search(text.strip())
    if not match:
        return None
    cleaned = match.group(0).replace("/", "-")
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d-%m-%y", "%m-%d-%Y"):
        try:
            return datetime.strptime(cleaned, fmt).date()
        except ValueError:
            continue
    return None


def extract_sales_from_sheet(sheet: Worksheet) -> dict[str, Decimal]:
    """Extracts daily PMS, AGO, and LPG sales volumes from summary block."""
    sales: dict[str, Decimal] = {}
    max_r = min(sheet.max_row or 30, 25)
    max_c = min(sheet.max_column or 35, 30)

    # Strategy: Scan for 'SALES' header column first for highest precision
    sales_header_col = None
    sales_header_row = None
    for r in range(1, max_r + 1):
        for c in range(10, max_c + 1):
            val = str(sheet.cell(r, c).value or "").strip().upper()
            if val == "SALES":
                # Check if rows directly beneath have PMS/AGO
                below_vals = {
                    str(sheet.cell(r + offset, c).value or "").strip().upper()
                    for offset in range(1, 4)
                }
                if below_vals & {"PMS", "AGO", "LPG"}:
                    sales_header_col = c
                    sales_header_row = r
                    break
        if sales_header_col:
            break

    if sales_header_col and sales_header_row:
        # Extract from the structured sales column
        for r in range(sales_header_row + 1, min(sales_header_row + 8, max_r + 1)):
            label = str(sheet.cell(r, sales_header_col).value or "").strip().upper()
            if label in {"PMS", "AGO", "LPG"} and label not in sales:
                num = parse_number(sheet.cell(r, sales_header_col + 1).value)
                if num is not None:
                    sales[label] = num
            elif label.startswith("TOTAL") or label in {"EXPENSES", "PRODUCTS RECEIVED"}:
                break

    # Secondary scan: If any product is missing, scan general grid for exact tokens
    for product in ("PMS", "AGO", "LPG"):
        if product in sales:
            continue
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                val = sheet.cell(r, c).value
                if isinstance(val, str) and val.strip().upper() == product:
                    # Check cell to right
                    num = parse_number(sheet.cell(r, c + 1).value)
                    if num is not None:
                        sales[product] = num
                        break
            if product in sales:
                break

    return sales


def extract_station_sales_records(
    workbook: Workbook,
    filename: str,
    config: SalesConfig,
    expected_month: int | None = None,
) -> list[DailySalesRecord]:
    """Extracts daily sales records from a station workbook based on config sync mode."""
    visible_sheets = [s for s in workbook.worksheets if s.sheet_state != "hidden"]
    if not visible_sheets:
        return []

    target_sheets: list[Worksheet] = []

    if config.sync_mode == "latest":
        try:
            target_sheets = [visible_sheets[config.source_sheet_index]]
        except IndexError:
            target_sheets = [visible_sheets[-1]]
    elif config.sync_mode == "specific_date":
        for s in visible_sheets:
            d = extract_date_from_sheet(
                s,
                filename=filename,
                target_date=config.target_date,
                expected_month=expected_month,
            )
            if d and d == config.target_date:
                target_sheets.append(s)
        if not target_sheets:
            # Try latest sheet as fallback or return empty
            return []
    else:  # "all"
        target_sheets = visible_sheets

    records: list[DailySalesRecord] = []
    for ws in target_sheets:
        d = extract_date_from_sheet(
            ws,
            filename=filename,
            target_date=config.target_date,
            expected_month=expected_month,
        )
        if not d:
            logger.warning("Could not resolve date for sheet %s in %s", ws.title, filename)
            continue

        if config.sync_mode == "specific_date" and d != config.target_date:
            continue

        sales_dict = extract_sales_from_sheet(ws)
        records.append(
            DailySalesRecord(
                station_file=filename,
                sheet_name=ws.title,
                date=d,
                pms=sales_dict.get("PMS"),
                ago=sales_dict.get("AGO"),
                lpg=sales_dict.get("LPG"),
            )
        )

    return records
