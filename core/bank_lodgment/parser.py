from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
import io
import json
import logging
from pathlib import Path
import re
from typing import Any

from core.gear.utils import parse_number

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class BankLodgmentRecord:
    message_id: str
    date: str
    time: str
    bank_officer_name: str
    bank_officer_bank: str
    amount: Decimal
    status: str
    manager_phone: str
    manager_handle: str
    raw_caption: str
    mapped_branch_id: str | None
    has_receipt_image: bool
    validation_flags: list[str] = field(default_factory=list)
    is_duplicate: bool = False


def normalize_iso_to_sheet_date(date_str: str) -> str:
    """Converts YYYY-MM-DD to DD-MM-YYYY, or preserves DD-MM-YYYY."""
    if not date_str:
        return ""
    date_str = date_str.strip()
    match = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", date_str)
    if match:
        year, month, day = match.groups()
        return f"{day}-{month}-{year}"
    return date_str


def parse_lodgment_json(source: Any) -> tuple[dict[str, Any], list[BankLodgmentRecord]]:
    """Parses a bank lodgment JSON export into structured records with duplicate detection."""
    if isinstance(source, (str, Path)) and Path(str(source)).exists():
        with open(source, "r", encoding="utf-8") as f:
            data = json.load(f)
    elif hasattr(source, "read"):
        content = source.read()
        if isinstance(content, bytes):
            content = content.decode("utf-8")
        data = json.loads(content)
    elif isinstance(source, (str, bytes)):
        if isinstance(source, bytes):
            source = source.decode("utf-8")
        data = json.loads(source)
    elif isinstance(source, dict):
        data = source
    else:
        raise ValueError(f"Unsupported JSON source type: {type(source)}")

    export_meta = data.get("export_meta", {})
    raw_records = data.get("records", [])

    records: list[BankLodgmentRecord] = []
    seen_confirmations: set[tuple[str, Decimal, str, str]] = set()

    for r in raw_records:
        msg_id = str(r.get("officer_message_id", "")).strip()
        rec_date = str(r.get("date", "")).strip()
        rec_time = str(r.get("time", "")).strip()

        bo = r.get("bank_officer", {}) or {}
        officer_name = str(bo.get("name", "")).strip()
        officer_bank = str(bo.get("bank", "")).strip()

        conf = r.get("confirmation", {}) or {}
        raw_amt = conf.get("parsed_amount")
        amount = parse_number(raw_amt) or Decimal("0")
        status = str(conf.get("status", "")).strip().upper()

        bs = r.get("branch_submission", {}) or {}
        phone = str(bs.get("manager_phone", "")).strip()
        handle = str(bs.get("manager_handle", "")).strip()
        caption = str(bs.get("raw_caption", "")).strip()
        branch_id = bs.get("mapped_branch_id")
        if branch_id is not None:
            branch_id = str(branch_id).strip()
        has_receipt = bool(bs.get("has_receipt_image", False))
        flags = [str(f) for f in r.get("validation_flags", [])]

        # Duplicate detection key
        dedup_key = (
            rec_date,
            amount,
            officer_name.upper(),
            branch_id or phone or handle or caption,
        )

        is_dup = False
        if status == "APPROVED":
            if dedup_key in seen_confirmations:
                is_dup = True
            else:
                seen_confirmations.add(dedup_key)

        records.append(
            BankLodgmentRecord(
                message_id=msg_id,
                date=rec_date,
                time=rec_time,
                bank_officer_name=officer_name,
                bank_officer_bank=officer_bank,
                amount=amount,
                status=status,
                manager_phone=phone,
                manager_handle=handle,
                raw_caption=caption,
                mapped_branch_id=branch_id,
                has_receipt_image=has_receipt,
                validation_flags=flags,
                is_duplicate=is_dup,
            )
        )

    return export_meta, records
