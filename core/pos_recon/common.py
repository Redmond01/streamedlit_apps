from __future__ import annotations

from datetime import datetime, timedelta
import re
from typing import Any


def cell_date_to_text(value: Any) -> str:
    """Converts a cell date value (datetime, Excel serial number, or string) to ISO YYYY-MM-DD string."""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, (int, float)):
        # Excel date serial (starts from 1899-12-30)
        try:
            return (datetime(1899, 12, 30) + timedelta(days=int(value))).date().isoformat()
        except Exception:
            pass
    if value is None:
        return ""
    return str(value).strip()


def normalize_pos(value: Any) -> str:
    """Normalizes POS terminal identifiers like 'SGR 02', 'obx-113', 'T.A 1' to 'SGR 2', 'OBX 113', 'TA 1'."""
    text = str(value or "").strip().upper().replace("-", " ")
    match = re.match(r"^(SGR|OBX|AGS|T\.?\s*A|TA)\s*(\d+)\b", text)
    if not match:
        return ""
    pos_type = match.group(1).replace(".", "").replace(" ", "")
    return f"{pos_type} {int(match.group(2))}"


def parse_money(value: Any) -> float | None:
    """Parses currency strings into float numbers, handling ₦, commas, negatives in parentheses."""
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return round(float(value), 2)
    text = str(value).strip()
    negative = text.startswith("-") or (text.startswith("(") and text.endswith(")"))
    text = (
        text.replace("₦", "")
        .replace(",", "")
        .replace("+", "")
        .replace("-", "")
        .replace("(", "")
        .replace(")", "")
        .strip()
    )
    if not text:
        return None
    try:
        number = round(float(text), 2)
        return -number if negative else number
    except ValueError:
        return None


def find_report_blocks(ws: Any) -> list[dict[str, Any]]:
    """Finds BANK STATEMENT report blocks in a reconciliation worksheet."""
    blocks = []
    max_c = ws.max_column or 50
    for col in range(1, max_c + 1):
        value = ws.cell(4, col).value
        if isinstance(value, str) and "bank statement" in value.lower():
            terminal_col = col - 2
            cashbook_col = col - 1
            bank_col = col
            difference_col = col + 1
            date = ""
            for date_col in range(max(1, terminal_col), min(max_c, difference_col) + 1):
                cell_val = ws.cell(1, date_col).value
                date = cell_date_to_text(cell_val) or date
            blocks.append({
                "date": date,
                "terminal_col": terminal_col,
                "cashbook_col": cashbook_col,
                "bank_col": bank_col,
                "difference_col": difference_col,
            })
    return blocks


def choose_block(blocks: list[dict[str, Any]], date: str | None = None, block_number: int | None = None) -> dict[str, Any]:
    """Selects the target report block based on date or 1-based block index."""
    if not blocks:
        raise ValueError("No report blocks found in the worksheet.")

    if block_number is not None:
        if block_number < 1 or block_number > len(blocks):
            raise ValueError(f"Block number {block_number} out of range (1..{len(blocks)}).")
        return blocks[block_number - 1]

    if date:
        matches = [b for b in blocks if b["date"] == date]
        if matches:
            return matches[0]
        if len(blocks) == 1:
            return blocks[0]
        available = ", ".join(b["date"] or f"block {i+1}" for i, b in enumerate(blocks))
        raise ValueError(f"No report block found for date '{date}'. Available blocks: {available}")

    if len(blocks) == 1:
        return blocks[0]

    raise ValueError(
        f"Workbook contains {len(blocks)} report blocks. Please specify a Date (YYYY-MM-DD) or Block Number."
    )
