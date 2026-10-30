from datetime import date, datetime, timedelta, timezone
import logging
from pathlib import Path
import tempfile
from typing import Any

import openpyxl
from openpyxl.styles import Font, PatternFill

from core.memory import cleanup_files, force_gc, spool_uploaded_file
from .common import (
    cell_date_to_text,
    choose_block,
    find_report_blocks,
    get_writable_cell,
    normalize_pos,
    parse_money,
)

logger = logging.getLogger(__name__)


def parse_timestamp(value: Any) -> datetime:
    """Parses a timestamp value from a cell (ISO string, datetime, date, formatted string).

    Returns a timezone-naive UTC datetime for consistent comparison across records.
    If parsing fails or value is empty, returns datetime.min.
    """
    if isinstance(value, datetime):
        if value.tzinfo is not None:
            return value.astimezone(timezone.utc).replace(tzinfo=None)
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if value is None:
        return datetime.min
    text = str(value).strip()
    if not text:
        return datetime.min
    try:
        dt = datetime.fromisoformat(text)
        if dt.tzinfo is not None:
            return dt.astimezone(timezone.utc).replace(tzinfo=None)
        return dt
    except Exception:
        pass
    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M:%S%z",
        "%d/%m/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
        "%Y/%m/%d %H:%M:%S",
        "%Y-%m-%d",
    ):
        try:
            dt = datetime.strptime(text, fmt)
            if dt.tzinfo is not None:
                return dt.astimezone(timezone.utc).replace(tzinfo=None)
            return dt
        except Exception:
            pass
    return datetime.min


def format_timestamp(value: Any) -> str:
    """Returns a readable string representation of a timestamp cell value."""
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return str(value or "").strip()


def daily_pos_sheet_name(date_str: str) -> str:
    """Returns sheet name formatted as Daily_POS_YYYY_MM_DD."""
    try:
        parsed = datetime.fromisoformat(date_str)
    except ValueError as error:
        raise ValueError("Date must use YYYY-MM-DD format (e.g. 2026-07-26).") from error
    return f"Daily_POS_{parsed:%Y_%m_%d}"


def choose_source_sheet(wb: Any, date_str: str, source_sheet: str | None = None) -> tuple[Any, str]:
    expected = source_sheet or daily_pos_sheet_name(date_str)
    if expected in wb.sheetnames:
        return wb[expected], expected

    available = ", ".join(wb.sheetnames)
    if source_sheet:
        raise ValueError(f"Source workbook does not contain sheet '{source_sheet}'. Available: {available}")

    raise ValueError(
        f"Source workbook does not contain expected sheet '{expected}' for date {date_str}. "
        f"Provide a sheet name override if needed. Available sheets: {available}"
    )


def find_header(headers: list[Any], wanted: str) -> int:
    wanted_clean = wanted.strip().lower()
    for index, value in enumerate(headers, start=1):
        if str(value or "").strip().lower() == wanted_clean:
            return index
    raise ValueError(f"Missing required source header: '{wanted}'")


def find_optional_header(headers: list[Any], wanted: str) -> int | None:
    wanted_clean = wanted.strip().lower()
    for index, value in enumerate(headers, start=1):
        if str(value or "").strip().lower() == wanted_clean:
            return index
    return None


def find_total_row(ws: Any, terminal_col: int) -> int:
    max_r = ws.max_row or 200
    for row in range(5, max_r + 1):
        if str(ws.cell(row, terminal_col).value or "").strip().upper() == "TOTAL":
            return row
    return max_r


def _extract_eod_legacy(
    ws: Any,
    selected_sheet: str,
    headers: list[Any],
    date_str: str,
    timestamp_col: int,
    submission_id_col: int,
    station_col: int,
    report_date_col: int,
    full_pos_col: int,
    amount_col: int,
    use_latest_duplicate: bool = True,
    remap_ags10: bool = True,
) -> dict[str, Any]:
    amount_by_pos: dict[str, float] = {}
    pos_submissions: dict[str, list[dict[str, Any]]] = {}
    kept_rows = []
    skipped_wrong_date = 0
    skipped_invalid_terminal = []
    skipped_invalid_amount = []
    source_rows = 0

    max_r = ws.max_row or 500
    for row in range(2, max_r + 1):
        row_values = [ws.cell(row, col).value for col in range(1, len(headers) + 1)]
        if all(value in (None, "") for value in row_values):
            continue
        source_rows += 1

        report_date = cell_date_to_text(ws.cell(row, report_date_col).value)
        if report_date != date_str:
            skipped_wrong_date += 1
            continue

        pos = normalize_pos(ws.cell(row, full_pos_col).value)
        if not pos:
            skipped_invalid_terminal.append({
                "row": row,
                "fullPosId": ws.cell(row, full_pos_col).value,
                "station": ws.cell(row, station_col).value,
            })
            continue

        if remap_ags10 and pos == "AGS 10":
            pos = "SGR 10"

        amount = parse_money(ws.cell(row, amount_col).value)
        if amount is None:
            skipped_invalid_amount.append({
                "row": row,
                "pos": pos,
                "amount": ws.cell(row, amount_col).value,
            })
            continue

        raw_timestamp = ws.cell(row, timestamp_col).value
        entry_info = {
            "sourceRow": row,
            "sourceSheet": selected_sheet,
            "timestamp": format_timestamp(raw_timestamp),
            "timestamp_raw": raw_timestamp,
            "submissionId": ws.cell(row, submission_id_col).value,
            "station": ws.cell(row, station_col).value,
            "sourcePos": ws.cell(row, full_pos_col).value,
            "pos": pos,
            "amount": amount,
            "entryMode": "LEGACY",
        }
        kept_rows.append(entry_info)

        if pos not in pos_submissions:
            pos_submissions[pos] = []
        pos_submissions[pos].append(entry_info)

    duplicate_terminals = []
    for pos, entries in pos_submissions.items():
        if len(entries) == 1:
            amount_by_pos[pos] = entries[0]["amount"]
        else:
            sorted_entries = sorted(
                entries,
                key=lambda e: (parse_timestamp(e["timestamp_raw"]), e["sourceRow"]),
            )
            latest_entry = sorted_entries[-1]
            superseded_entries = sorted_entries[:-1]

            if use_latest_duplicate:
                computed_amount = latest_entry["amount"]
            else:
                computed_amount = round(sum(e["amount"] for e in entries), 2)

            amount_by_pos[pos] = computed_amount

            duplicate_terminals.append({
                "pos": pos,
                "count": len(entries),
                "amounts": [e["amount"] for e in sorted_entries],
                "totalSum": round(sum(e["amount"] for e in entries), 2),
                "computedAmount": computed_amount,
                "latestSubmission": latest_entry,
                "latestTimestamp": latest_entry["timestamp"],
                "latestSubmissionId": latest_entry["submissionId"],
                "latestStation": latest_entry["station"],
                "supersededSubmissions": superseded_entries,
                "submissions": sorted_entries,
                "isCrossDayDuplicate": False,
            })

    return {
        "sourceRows": source_rows,
        "keptRows": kept_rows,
        "skippedWrongDate": skipped_wrong_date,
        "skippedInvalidTerminal": skipped_invalid_terminal,
        "skippedInvalidAmount": skipped_invalid_amount,
        "amountByPos": amount_by_pos,
        "duplicateTerminals": duplicate_terminals,
        "useLatestDuplicate": use_latest_duplicate,
        "sumDuplicates": not use_latest_duplicate,
        "remapAgs10": remap_ags10,
        "sourceSheet": selected_sheet,
        "entryModeEnabled": False,
        "prevSourceSheet": None,
        "defaultCount": len(amount_by_pos),
        "currentFillinCount": 0,
        "crossDayDuplicatesCount": 0,
        "crossDayDuplicates": [],
        "fillinTerminals": [],
    }


def _extract_eod_entry_mode(
    wb: Any,
    ws: Any,
    selected_sheet: str,
    headers: list[Any],
    date_str: str,
    timestamp_col: int,
    submission_id_col: int,
    station_col: int,
    report_date_col: int,
    full_pos_col: int,
    amount_col: int,
    entry_mode_col: int,
    use_latest_duplicate: bool = True,
    remap_ags10: bool = True,
) -> dict[str, Any]:
    kept_rows = []
    skipped_wrong_date = 0
    skipped_invalid_terminal = []
    skipped_invalid_amount = []
    source_rows = 0

    default_submissions_by_pos: dict[str, list[dict[str, Any]]] = {}

    max_r = ws.max_row or 500
    for row in range(2, max_r + 1):
        row_values = [ws.cell(row, col).value for col in range(1, len(headers) + 1)]
        if all(value in (None, "") for value in row_values):
            continue
        source_rows += 1

        mode = str(ws.cell(row, entry_mode_col).value or "").strip().upper()
        if mode != "DEFAULT":
            continue

        report_date = cell_date_to_text(ws.cell(row, report_date_col).value)
        if report_date and report_date != date_str:
            skipped_wrong_date += 1
            continue

        pos = normalize_pos(ws.cell(row, full_pos_col).value)
        if not pos:
            skipped_invalid_terminal.append({
                "row": row,
                "fullPosId": ws.cell(row, full_pos_col).value,
                "station": ws.cell(row, station_col).value,
            })
            continue

        if remap_ags10 and pos == "AGS 10":
            pos = "SGR 10"

        amount = parse_money(ws.cell(row, amount_col).value)
        if amount is None:
            skipped_invalid_amount.append({
                "row": row,
                "pos": pos,
                "amount": ws.cell(row, amount_col).value,
            })
            continue

        raw_timestamp = ws.cell(row, timestamp_col).value
        entry_info = {
            "sourceRow": row,
            "sourceSheet": selected_sheet,
            "timestamp": format_timestamp(raw_timestamp),
            "timestamp_raw": raw_timestamp,
            "submissionId": ws.cell(row, submission_id_col).value,
            "station": ws.cell(row, station_col).value,
            "sourcePos": ws.cell(row, full_pos_col).value,
            "pos": pos,
            "amount": amount,
            "entryMode": "DEFAULT",
        }
        default_submissions_by_pos.setdefault(pos, []).append(entry_info)

    # 1. Resolve DEFAULT entries (and any intra-mode duplicate in DEFAULT)
    default_entries_by_pos: dict[str, dict[str, Any]] = {}
    amount_by_pos: dict[str, float] = {}

    for pos, entries in default_submissions_by_pos.items():
        if len(entries) == 1:
            chosen = entries[0]
        else:
            sorted_entries = sorted(
                entries,
                key=lambda e: (parse_timestamp(e["timestamp_raw"]), e["sourceRow"]),
            )
            chosen = sorted_entries[-1]
        default_entries_by_pos[pos] = chosen
        amount_by_pos[pos] = chosen["amount"]
        kept_rows.append(chosen)

    # 2. Look up Previous Day sheet for CURRENT submissions
    parsed_date = datetime.fromisoformat(date_str).date()
    prev_date_str = (parsed_date - timedelta(days=1)).isoformat()
    prev_sheet_name = daily_pos_sheet_name(prev_date_str)

    current_submissions_by_pos: dict[str, list[dict[str, Any]]] = {}
    prev_sheet_found = False

    if prev_sheet_name in wb.sheetnames:
        pws = wb[prev_sheet_name]
        p_headers = [cell.value for cell in pws[1]]
        p_entry_mode_col = find_optional_header(p_headers, "Entry Mode")

        if p_entry_mode_col is not None:
            prev_sheet_found = True
            p_timestamp_col = find_header(p_headers, "Timestamp")
            p_sub_col = find_header(p_headers, "Submission ID")
            p_station_col = find_header(p_headers, "Station")
            p_pos_col = find_header(p_headers, "Full POS ID")
            p_amount_col = find_header(p_headers, "Amount")

            p_max_r = pws.max_row or 500
            for row in range(2, p_max_r + 1):
                p_values = [pws.cell(row, col).value for col in range(1, len(p_headers) + 1)]
                if all(v in (None, "") for v in p_values):
                    continue

                p_mode = str(pws.cell(row, p_entry_mode_col).value or "").strip().upper()
                if p_mode != "CURRENT":
                    continue

                pos = normalize_pos(pws.cell(row, p_pos_col).value)
                if not pos:
                    continue

                if remap_ags10 and pos == "AGS 10":
                    pos = "SGR 10"

                amount = parse_money(pws.cell(row, p_amount_col).value)
                if amount is None:
                    continue

                raw_ts = pws.cell(row, p_timestamp_col).value
                p_entry_info = {
                    "sourceRow": row,
                    "sourceSheet": prev_sheet_name,
                    "timestamp": format_timestamp(raw_ts),
                    "timestamp_raw": raw_ts,
                    "submissionId": pws.cell(row, p_sub_col).value,
                    "station": pws.cell(row, p_station_col).value,
                    "sourcePos": pws.cell(row, p_pos_col).value,
                    "pos": pos,
                    "amount": amount,
                    "entryMode": "CURRENT",
                }
                current_submissions_by_pos.setdefault(pos, []).append(p_entry_info)

    # Resolve any intra-mode duplicate in CURRENT
    current_entries_by_pos: dict[str, dict[str, Any]] = {}
    for pos, entries in current_submissions_by_pos.items():
        if len(entries) == 1:
            current_entries_by_pos[pos] = entries[0]
        else:
            sorted_entries = sorted(
                entries,
                key=lambda e: (parse_timestamp(e["timestamp_raw"]), e["sourceRow"]),
            )
            current_entries_by_pos[pos] = sorted_entries[-1]

    # 3. Cross-Day Evaluation
    cross_day_duplicates = []
    fillin_terminals = []
    duplicate_terminals = []

    for pos, c_entry in current_entries_by_pos.items():
        if pos in default_entries_by_pos:
            # Duplicate found across days!
            # New directive: today's DEFAULT is TRUE AMOUNT, yesterday's CURRENT is superseded!
            d_entry = default_entries_by_pos[pos]
            if use_latest_duplicate:
                chosen_amount = d_entry["amount"]
            else:
                chosen_amount = round(d_entry["amount"] + c_entry["amount"], 2)
            amount_by_pos[pos] = chosen_amount

            dup_item = {
                "pos": pos,
                "count": 2,
                "amounts": [d_entry["amount"], c_entry["amount"]],
                "totalSum": round(d_entry["amount"] + c_entry["amount"], 2),
                "computedAmount": chosen_amount,
                "defaultAmount": d_entry["amount"],
                "currentAmount": c_entry["amount"],
                "defaultTimestamp": d_entry["timestamp"],
                "currentTimestamp": c_entry["timestamp"],
                "defaultSubmissionId": d_entry["submissionId"],
                "currentSubmissionId": c_entry["submissionId"],
                "defaultStation": d_entry["station"],
                "currentStation": c_entry["station"],
                "chosenMode": "DEFAULT" if use_latest_duplicate else "SUM",
                "chosenAmount": chosen_amount,
                "latestTimestamp": d_entry["timestamp"],
                "latestSubmissionId": d_entry["submissionId"],
                "isCrossDayDuplicate": True,
                "priorSheet": prev_sheet_name,
            }
            cross_day_duplicates.append(dup_item)
            duplicate_terminals.append(dup_item)
        else:
            # Terminal was NOT submitted today as DEFAULT!
            # Pick yesterday's CURRENT as fill-in!
            amount_by_pos[pos] = c_entry["amount"]
            fillin_terminals.append(c_entry)
            kept_rows.append(c_entry)

    return {
        "sourceRows": source_rows,
        "keptRows": kept_rows,
        "skippedWrongDate": skipped_wrong_date,
        "skippedInvalidTerminal": skipped_invalid_terminal,
        "skippedInvalidAmount": skipped_invalid_amount,
        "amountByPos": amount_by_pos,
        "duplicateTerminals": duplicate_terminals,
        "useLatestDuplicate": use_latest_duplicate,
        "sumDuplicates": not use_latest_duplicate,
        "remapAgs10": remap_ags10,
        "sourceSheet": selected_sheet,
        "entryModeEnabled": True,
        "prevSourceSheet": prev_sheet_name if prev_sheet_found else None,
        "defaultCount": len(default_entries_by_pos),
        "currentFillinCount": len(fillin_terminals),
        "crossDayDuplicatesCount": len(cross_day_duplicates),
        "crossDayDuplicates": cross_day_duplicates,
        "fillinTerminals": fillin_terminals,
    }


def extract_eod_source(
    source_path: Path,
    date_str: str,
    source_sheet: str | None = None,
    use_latest_duplicate: bool = True,
    sum_duplicates: bool | None = None,
    remap_ags10: bool = True,
) -> dict[str, Any]:
    if sum_duplicates is not None:
        use_latest_duplicate = not sum_duplicates

    wb = openpyxl.load_workbook(source_path, data_only=True)
    try:
        ws, selected_sheet = choose_source_sheet(wb, date_str, source_sheet=source_sheet)

        headers = [cell.value for cell in ws[1]]
        timestamp_col = find_header(headers, "Timestamp")
        submission_id_col = find_header(headers, "Submission ID")
        station_col = find_header(headers, "Station")
        report_date_col = find_header(headers, "Report Date")
        full_pos_col = find_header(headers, "Full POS ID")
        amount_col = find_header(headers, "Amount")
        entry_mode_col = find_optional_header(headers, "Entry Mode")

        if entry_mode_col is not None:
            return _extract_eod_entry_mode(
                wb=wb,
                ws=ws,
                selected_sheet=selected_sheet,
                headers=headers,
                date_str=date_str,
                timestamp_col=timestamp_col,
                submission_id_col=submission_id_col,
                station_col=station_col,
                report_date_col=report_date_col,
                full_pos_col=full_pos_col,
                amount_col=amount_col,
                entry_mode_col=entry_mode_col,
                use_latest_duplicate=use_latest_duplicate,
                remap_ags10=remap_ags10,
            )

        return _extract_eod_legacy(
            ws=ws,
            selected_sheet=selected_sheet,
            headers=headers,
            date_str=date_str,
            timestamp_col=timestamp_col,
            submission_id_col=submission_id_col,
            station_col=station_col,
            report_date_col=report_date_col,
            full_pos_col=full_pos_col,
            amount_col=amount_col,
            use_latest_duplicate=use_latest_duplicate,
            remap_ags10=remap_ags10,
        )
    finally:
        wb.close()


def write_eod_to_workbook(
    source_input: Any,
    target_workbook_input: Any,
    date_str: str,
    source_sheet: str | None = None,
    block_number: int | None = None,
    use_latest_duplicate: bool = True,
    sum_duplicates: bool | None = None,
    remap_ags10: bool = True,
) -> tuple[Path, dict[str, Any]]:
    """Syncs Moniepoint EOD transactions into reconciliation workbook's CASHBOOK column."""
    if not date_str:
        raise ValueError("Report Date is required to filter EOD records safely.")

    if sum_duplicates is not None:
        use_latest_duplicate = not sum_duplicates

    src_tmp = spool_uploaded_file(source_input, prefix="eod_src_")
    tgt_tmp = spool_uploaded_file(target_workbook_input, prefix="recon_tgt_")

    target_wb = None
    output_path: Path | None = None

    try:
        # Extract EOD data from source file
        extracted = extract_eod_source(
            src_tmp,
            date_str,
            source_sheet=source_sheet,
            use_latest_duplicate=use_latest_duplicate,
            remap_ags10=remap_ags10,
        )

        # Open target workbook for modification
        target_wb = openpyxl.load_workbook(tgt_tmp)
        ws = target_wb.active
        blocks = find_report_blocks(ws)
        if not blocks:
            raise ValueError("No BANK STATEMENT report blocks found in target workbook.")

        block = choose_block(blocks, date=date_str, block_number=block_number)

        total_row = find_total_row(ws, block["terminal_col"])
        data_end_row = total_row - 1
        unmatched = set(extracted["amountByPos"])
        mapped_rows = []

        try:
            date_cell = get_writable_cell(ws, 1, block["bank_col"])
            try:
                date_cell.value = datetime.fromisoformat(date_str)
                date_cell.number_format = "yyyy-mm-dd"
            except Exception:
                date_cell.value = date_str
        except Exception as exc:
            logger.warning("Could not update cosmetic date in row 1 for block: %s", exc)

        for row in range(5, data_end_row + 1):
            cb_cell = get_writable_cell(ws, row, block["cashbook_col"])
            bk_cell = get_writable_cell(ws, row, block["bank_col"])
            df_cell = get_writable_cell(ws, row, block["difference_col"])

            cb_cell.value = None
            bk_cell.value = None
            df_cell.value = (
                f"=SUM({cb_cell.coordinate}-"
                f"{bk_cell.coordinate})"
            )

            pos = normalize_pos(ws.cell(row, block["terminal_col"]).value)
            if not pos or pos not in extracted["amountByPos"]:
                continue

            amount = extracted["amountByPos"][pos]
            cb_cell.value = amount
            cb_cell.number_format = "#,##0.00"
            mapped_rows.append({
                "row": row,
                "workbookPos": ws.cell(row, block["terminal_col"]).value,
                "pos": pos,
                "amount": amount,
            })
            unmatched.discard(pos)

        tot_cb = get_writable_cell(ws, total_row, block["cashbook_col"])
        tot_bk = get_writable_cell(ws, total_row, block["bank_col"])
        tot_df = get_writable_cell(ws, total_row, block["difference_col"])

        tot_cb.value = f"=SUM({ws.cell(5, block['cashbook_col']).coordinate}:{ws.cell(data_end_row, block['cashbook_col']).coordinate})"
        tot_bk.value = f"=SUM({ws.cell(5, block['bank_col']).coordinate}:{ws.cell(data_end_row, block['bank_col']).coordinate})"
        tot_df.value = (
            f"=SUM({tot_cb.coordinate}-"
            f"{tot_bk.coordinate})"
        )

        _write_audit_sheet(
            target_wb,
            source_filename=getattr(source_input, "name", "EOD_Source.xlsx"),
            target_filename=getattr(target_workbook_input, "name", "Reconciliation.xlsx"),
            date_str=date_str,
            extracted=extracted,
            mapped_rows=mapped_rows,
            unmatched=sorted(unmatched),
        )

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", prefix=f"EOD_SYNC_{date_str.replace('-', '')}_") as out_tmp:
            output_path = Path(out_tmp.name)
        target_wb.save(output_path)

        result_stats = {
            "sourceSheet": extracted["sourceSheet"],
            "sourceRows": extracted["sourceRows"],
            "keptRows": len(extracted["keptRows"]),
            "skippedWrongDate": extracted["skippedWrongDate"],
            "skippedInvalidTerminal": len(extracted["skippedInvalidTerminal"]),
            "skippedInvalidAmount": len(extracted["skippedInvalidAmount"]),
            "uniqueTerminals": len(extracted["amountByPos"]),
            "mappedRows": len(mapped_rows),
            "unmatchedSourceTerminals": sorted(unmatched),
            "duplicateTerminals": extracted.get("duplicateTerminals", []),
            "useLatestDuplicate": use_latest_duplicate,
            "sumDuplicates": not use_latest_duplicate,
            "remapAgs10": extracted.get("remapAgs10", remap_ags10),
            "entryModeEnabled": extracted.get("entryModeEnabled", False),
            "prevSourceSheet": extracted.get("prevSourceSheet"),
            "defaultCount": extracted.get("defaultCount", 0),
            "currentFillinCount": extracted.get("currentFillinCount", 0),
            "crossDayDuplicatesCount": extracted.get("crossDayDuplicatesCount", 0),
            "crossDayDuplicates": extracted.get("crossDayDuplicates", []),
            "fillinTerminals": extracted.get("fillinTerminals", []),
        }
        return output_path, result_stats

    finally:
        if target_wb:
            try:
                target_wb.close()
            except Exception:
                pass
        cleanup_files(src_tmp, tgt_tmp)
        force_gc()


def _write_audit_sheet(
    wb: Any,
    source_filename: str,
    target_filename: str,
    date_str: str,
    extracted: dict[str, Any],
    mapped_rows: list[dict[str, Any]],
    unmatched: list[str],
) -> None:
    audit_sheet_name = "EOD Extract Audit"
    if audit_sheet_name in wb.sheetnames:
        del wb[audit_sheet_name]
    ws = wb.create_sheet(audit_sheet_name)

    entry_mode_enabled = extracted.get("entryModeEnabled", False)
    remap_ags10 = extracted.get("remapAgs10", True)

    if entry_mode_enabled:
        cross_dups = extracted.get("crossDayDuplicates", [])
        fillins = extracted.get("fillinTerminals", [])

        rows: list[list[Any]] = [
            ["Metric", "Value"],
            ["Source Workbook", source_filename],
            ["Target Date Sheet (DEFAULT)", extracted.get("sourceSheet", "")],
            ["Prior Date Sheet (CURRENT)", extracted.get("prevSourceSheet") or "None (Not present)"],
            ["Target Workbook", target_filename],
            ["Target Date", date_str],
            ["Mapped Records into Cashbook", len(mapped_rows)],
            ["Target Date (DEFAULT) Submissions", extracted.get("defaultCount", 0)],
            ["Prior Day (CURRENT) Fill-ins Adopted", extracted.get("currentFillinCount", 0)],
            ["Cross-Day Duplicates Resolved (DEFAULT Kept)", extracted.get("crossDayDuplicatesCount", 0)],
            ["Remap AGS 10 to SGR 10", "Enabled" if remap_ags10 else "Disabled"],
            ["Unmatched Source Terminals", len(unmatched)],
            ["Total Rows Kept", len(extracted["keptRows"])],
            [],
        ]

        if cross_dups:
            rows.extend([
                ["Cross-Day Duplicates (Today DEFAULT Kept as TRUE Amount, Prior CURRENT Discarded)"],
                [
                    "POS Terminal",
                    "Selected TRUE Amount (DEFAULT)",
                    "Today Submission ID",
                    "Today Timestamp",
                    "Today Station",
                    "Discarded Amount (Prior CURRENT)",
                    "Prior Submission ID",
                    "Prior Timestamp",
                    "Prior Station",
                ],
            ])
            for d in cross_dups:
                rows.append([
                    d["pos"],
                    d["defaultAmount"],
                    d.get("defaultSubmissionId", "-"),
                    d.get("defaultTimestamp", "-"),
                    d.get("defaultStation", "-"),
                    d["currentAmount"],
                    d.get("currentSubmissionId", "-"),
                    d.get("currentTimestamp", "-"),
                    d.get("currentStation", "-"),
                ])
            rows.append([])

        if fillins:
            rows.extend([
                ["Prior Day CURRENT Fill-ins (No DEFAULT Submission Today)"],
                [
                    "POS Terminal",
                    "Adopted Amount",
                    "Timestamp",
                    "Submission ID",
                    "Station",
                    "Source Sheet",
                ],
            ])
            for f in fillins:
                rows.append([
                    f["pos"],
                    f["amount"],
                    f.get("timestamp", "-"),
                    f.get("submissionId", "-"),
                    f.get("station", "-"),
                    f.get("sourceSheet", "-"),
                ])
            rows.append([])

        rows.extend([
            ["Mapped Row", "Workbook Terminal", "Normalized Terminal", "EOD Amount"],
        ])
        rows.extend([r["row"], r["workbookPos"], r["pos"], r["amount"]] for r in mapped_rows)
        rows.extend([[], ["Unmatched Source Terminal"]])
        rows.extend([terminal] for terminal in unmatched)
        rows.extend([[], ["Source Row", "Source Sheet", "Entry Mode", "Timestamp", "Submission ID", "Station", "Source POS", "Normalized POS", "Amount"]])
        rows.extend([
            r["sourceRow"],
            r.get("sourceSheet", "-"),
            r.get("entryMode", "-"),
            r["timestamp"],
            r["submissionId"],
            r["station"],
            r["sourcePos"],
            r["pos"],
            r["amount"],
        ] for r in extracted["keptRows"])

    else:
        duplicate_list = extracted.get("duplicateTerminals", [])
        use_latest = extracted.get("useLatestDuplicate", True)

        rows = [
            ["Metric", "Value"],
            ["Source Workbook", source_filename],
            ["Source Sheet", extracted.get("sourceSheet", "")],
            ["Target Workbook", target_filename],
            ["Target Date", date_str],
            ["Mapped Records", len(mapped_rows)],
            ["Duplicate Terminals Count", len(duplicate_list)],
            [
                "Duplicate Handling Mode",
                "Latest Submission by Timestamp (Default)" if use_latest else "Sum Amounts (Disabled)",
            ],
            ["Remap AGS 10 to SGR 10", "Enabled" if remap_ags10 else "Disabled"],
            ["Unmatched Source Terminals", len(unmatched)],
            ["Rows Kept for Target Date", len(extracted["keptRows"])],
            ["Rows Skipped for Other Dates", extracted["skippedWrongDate"]],
            [],
        ]

        if duplicate_list:
            dup_title = (
                "Duplicate POS Terminals (Resolved by Latest Timestamp)"
                if use_latest
                else "Duplicate POS Terminals (Summed Together)"
            )
            rows.extend([
                [dup_title],
                [
                    "POS Terminal",
                    "Submission Count",
                    "Latest Timestamp",
                    "Latest Submission ID",
                    "Selected True Amount",
                    "All Submitted Amounts",
                    "Total Sum If Summed",
                ],
            ])
            for dup in duplicate_list:
                rows.append([
                    dup["pos"],
                    dup["count"],
                    dup.get("latestTimestamp", "-"),
                    dup.get("latestSubmissionId", "-"),
                    dup["computedAmount"],
                    ", ".join(f"₦{a:,.2f}" for a in dup["amounts"]),
                    dup["totalSum"],
                ])
            rows.append([])

        rows.extend([
            ["Mapped Row", "Workbook Terminal", "Normalized Terminal", "EOD Amount"],
        ])
        rows.extend([r["row"], r["workbookPos"], r["pos"], r["amount"]] for r in mapped_rows)
        rows.extend([[], ["Unmatched Source Terminal"]])
        rows.extend([terminal] for terminal in unmatched)
        rows.extend([[], ["Kept Source Row", "Timestamp", "Submission ID", "Station", "Source POS", "Normalized POS", "Amount"]])
        rows.extend([
            r["sourceRow"],
            r["timestamp"],
            r["submissionId"],
            r["station"],
            r["sourcePos"],
            r["pos"],
            r["amount"],
        ] for r in extracted["keptRows"])

    for row_idx, row_values in enumerate(rows, start=1):
        for col_idx, val in enumerate(row_values, start=1):
            ws.cell(row_idx, col_idx).value = val

    header_font = Font(bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    for cell in ws[10]:
        cell.font = header_font
        cell.fill = PatternFill(fill_type="solid", fgColor="5B9BD5")

    for col in ws.columns:
        max_length = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(max_length + 2, 12), 45)
