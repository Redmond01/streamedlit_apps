from __future__ import annotations

from datetime import datetime
import re
from typing import Any

from core.gear.matcher import _normalize_station_name, _station_match_score
from .config import BankLodgmentConfig
from .parser import BankLodgmentRecord


def format_lodgment_time(time_str: str) -> str:
    """Formats 24-hour time '14:36' into Excel standard '2.36PM' or '09:46' into '9.46AM'."""
    if not time_str:
        return ""
    time_str = time_str.strip()
    match = re.search(r"(\d{1,2})[:.](\d{2})", time_str)
    if not match:
        return time_str

    hour, minute = int(match.group(1)), match.group(2)
    meridiem = "AM" if hour < 12 else "PM"
    hour_12 = hour % 12
    if hour_12 == 0:
        hour_12 = 12

    return f"{hour_12}.{minute}{meridiem}"


def normalize_bank_name(bank_str: str) -> str:
    """Normalizes bank officer bank string to workbook standard."""
    if not bank_str:
        return "UNKNOWN"
    text = bank_str.strip().upper()
    if "UNION" in text:
        return "UNION"
    if "FCMB" in text:
        return "FCMB"
    if "ACCESS" in text:
        return "FCMB"
    if "ZENITH" in text:
        return "ZENITH"
    if "PROVIDUS" in text:
        return "PROVIDUS"
    return text


def resolve_station_for_record(
    record: BankLodgmentRecord,
    config: BankLodgmentConfig,
    known_stations: list[str],
    depositors_map: dict[str, str] | None = None,
) -> tuple[str | None, str, float]:
    """Resolves target station name using a multi-signal waterfall.
    
    Returns:
        (station_name, resolution_strategy, confidence_score)
    """
    # 1. Direct Mapped Branch ID
    if record.mapped_branch_id and record.mapped_branch_id in config.branch_id_map:
        return config.branch_id_map[record.mapped_branch_id], "branch_id", 1.0

    # 2. Manager Phone Directory
    clean_phone = re.sub(r"[^\d+]", "", record.manager_phone)
    if clean_phone and clean_phone in config.phone_map:
        return config.phone_map[clean_phone], "manager_phone", 0.98

    # 3. Manager Handle Directory
    if record.manager_handle and record.manager_handle in config.handle_map:
        return config.handle_map[record.manager_handle], "manager_handle", 0.95

    # 4. Raw Caption Analysis
    caption = record.raw_caption.strip() if record.raw_caption else ""
    if caption and caption.upper() not in {"PHOTO", "DOCUMENT", "IMAGE"}:
        cap_upper = caption.upper()

        # Known domain aliases
        if "CAR GO 2" in cap_upper or "CAR-GO 2" in cap_upper or "CARGO 2" in cap_upper:
            return "Cargo 2 OBX", "caption_alias", 0.95
        if "CAR-GO 1" in cap_upper or "CAR GO 1" in cap_upper or "CARGO 1" in cap_upper:
            return "Cargo 1 OBX", "caption_alias", 0.95
        if "OYO 1" in cap_upper:
            return "Oyo Ilora", "caption_alias", 0.95
        if "AKODO" in cap_upper:
            return "Eleko 2", "caption_alias", 0.95
        if "AGODI GATE" in cap_upper:
            return "Agodi OBX", "caption_alias", 0.95
        if "LEKKI 4" in cap_upper:
            return "Ajah", "caption_alias", 0.95
        if "MONATAN" in cap_upper:
            return "Iwo 2 Monatan OBX", "caption_alias", 0.95
        if "IFEROAD1" in cap_upper or "IFE ROAD 1" in cap_upper or "IFE 1" in cap_upper:
            return "Ife 1 OBX", "caption_alias", 0.95
        if "IFE ROAD 3" in cap_upper or "IFE 3" in cap_upper:
            return "Ife 3 OBX", "caption_alias", 0.95

        # Scored fuzzy token match across known stations
        norm_cap = _normalize_station_name(caption)
        scored: list[tuple[float, str]] = []
        for cand in known_stations:
            norm_cand = _normalize_station_name(cand)
            score = _station_match_score(norm_cap, norm_cand)
            scored.append((score, cand))

        scored.sort(key=lambda item: item[0], reverse=True)
        if scored and scored[0][0] >= 0.50:
            return scored[0][1], "caption_fuzzy", scored[0][0]

    # 5. Depositor Name matching (fallback)
    if depositors_map and record.manager_handle:
        norm_handle = re.sub(r"[^A-Z ]", "", record.manager_handle.upper()).strip()
        for dep_name, st_name in depositors_map.items():
            norm_dep = re.sub(r"[^A-Z ]", "", dep_name.upper()).strip()
            if norm_handle and norm_dep:
                tokens_handle = set(norm_handle.split())
                tokens_dep = set(norm_dep.split())
                if len(tokens_handle & tokens_dep) >= 2:
                    return st_name, "depositor_match", 0.90

    return None, "unmatched", 0.0
