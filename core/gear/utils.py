from __future__ import annotations

from decimal import Decimal, InvalidOperation
import re
from typing import Any


TANK_PATTERN = re.compile(r"\b(PMS|AGO)\s*0*(\d*)\b", re.IGNORECASE)


def safe_cell_value(value: Any) -> Any:
    if isinstance(value, str):
        return value.strip()
    return value


def is_blank(value: Any) -> bool:
    value = safe_cell_value(value)
    return value is None or value == ""


def normalize_tank_name(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip().upper().replace("AG0", "AGO")
    match = TANK_PATTERN.search(text)
    if not match:
        return None
    number = int(match.group(2) or "1")
    return f"{match.group(1).upper()}{number}"


def is_valid_tank(value: Any) -> bool:
    normalized = normalize_tank_name(value)
    return normalized is not None and (normalized.startswith("PMS") or normalized.startswith("AGO"))


def normalize_header(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().upper())


def parse_number(value: Any) -> Decimal | None:
    if is_blank(value):
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    text = str(value).strip()
    if text.startswith("="):
        return None
    if re.search(r"[A-Za-z]", text):
        return None
    text = text.replace(",", "")
    if text.endswith("%"):
        text = text[:-1].strip()
    text = re.sub(r"[^0-9.\-]", "", text)
    if text in {"", "-", ".", "-."}:
        return None
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def decimal_to_excel(value: Decimal | None) -> int | float | None:
    if value is None:
        return None
    if value == value.to_integral_value():
        return int(value)
    return float(value)
