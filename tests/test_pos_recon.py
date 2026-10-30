from __future__ import annotations

from pathlib import Path
from core.pos_recon.common import normalize_pos, parse_money, cell_date_to_text
from core.pos_recon.cashbook_extractor import extract_cashbook_loop
from core.pos_recon.bank_loader import load_json_to_workbook
from core.memory import cleanup_files


BANK_DIR = Path("/home/im-redmond/Desktop/Desktop/bank_statement/scripts")


def test_pos_normalization():
    assert normalize_pos("SGR 02") == "SGR 2"
    assert normalize_pos("OBX 113") == "OBX 113"
    assert normalize_pos("T.A 1") == "TA 1"
    assert normalize_pos("TA 04") == "TA 4"
    assert normalize_pos("AGS 005") == "AGS 5"
    assert normalize_pos("UNKNOWN 1") == ""


def test_parse_money():
    assert parse_money("₦12,345.67") == 12345.67
    assert parse_money("(500.00)") == -500.00
    assert parse_money("-250.50") == -250.50
    assert parse_money(" 1,000 ") == 1000.0
    assert parse_money("") is None
    assert parse_money(None) is None


def test_cashbook_extractor_with_real_file():
    sample_file = BANK_DIR / "Book22SEPT.xlsx"
    if not sample_file.exists():
        return

    result = extract_cashbook_loop(
        workbook_input=sample_file,
        date="2026-09-22",
        min_number=11,
        include_audit=True,
    )
    assert result["date"] == "2026-09-22"
    assert "groups" in result
    assert "SGR" in result["groups"]
    assert "OBX" in result["groups"]
    assert len(result["groups"]["SGR"]) > 0


def test_bank_loader_with_real_file():
    sample_wb = BANK_DIR / "Book22SEPT.xlsx"
    sample_json = BANK_DIR / "pos_indexeddb_export_day_22.json"
    if not sample_wb.exists() or not sample_json.exists():
        return

    out_path, stats = load_json_to_workbook(
        workbook_input=sample_wb,
        json_input=sample_json,
        date="2026-09-22",
    )
    try:
        assert out_path.exists()
        assert stats["matched"] > 0
        assert stats["okWritten"] > 0
    finally:
        cleanup_files(out_path)


def test_eod_sync_integration():
    import tempfile
    import openpyxl
    from core.pos_recon.eod_extractor import write_eod_to_workbook

    # Create mock EOD source file
    eod_wb = openpyxl.Workbook()
    ews = eod_wb.active
    ews.title = "Daily_POS_2026_09_22"

    headers = ["Timestamp", "Submission ID", "Station", "Report Date", "Full POS ID", "Amount"]
    ews.append(headers)
    ews.append(["2026-09-22 10:00:00", "SUB1", "Station 1", "2026-09-22", "SGR 12", 15000])
    ews.append(["2026-09-22 11:00:00", "SUB2", "Station 2", "2026-09-22", "OBX 113", 25000])

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f_eod:
        eod_wb.save(f_eod.name)
        eod_path = Path(f_eod.name)

    sample_wb = BANK_DIR / "Book22SEPT.xlsx"
    if not sample_wb.exists():
        cleanup_files(eod_path)
        return

    try:
        out_path, stats = write_eod_to_workbook(
            source_input=eod_path,
            target_workbook_input=sample_wb,
            date_str="2026-09-22",
        )
        assert out_path.exists()
        assert stats["sourceRows"] == 2
        assert stats["keptRows"] == 2
        assert stats["mappedRows"] > 0

        # Verify audit sheet was written
        wb_check = openpyxl.load_workbook(out_path)
        assert "EOD Extract Audit" in wb_check.sheetnames
        wb_check.close()
        cleanup_files(out_path)

    finally:
        cleanup_files(eod_path)


def test_get_writable_cell_with_merged_range():
    import openpyxl
    from openpyxl.cell.cell import Cell, MergedCell
    from core.pos_recon.common import get_writable_cell

    wb = openpyxl.Workbook()
    ws = wb.active
    # Merge B1:D2
    ws.merge_cells("B1:D2")
    ws["B1"] = "Master Date"

    # B1 is normal Cell
    cell_b1 = get_writable_cell(ws, 1, 2)
    assert isinstance(cell_b1, Cell)
    assert not isinstance(cell_b1, MergedCell)
    assert cell_b1.coordinate == "B1"

    # C1 is a MergedCell, must resolve to master cell B1
    assert isinstance(ws.cell(1, 3), MergedCell)
    cell_c1 = get_writable_cell(ws, 1, 3)
    assert isinstance(cell_c1, Cell)
    assert not isinstance(cell_c1, MergedCell)
    assert cell_c1.coordinate == "B1"

    # D2 is a MergedCell, must also resolve to master cell B1
    assert isinstance(ws.cell(2, 4), MergedCell)
    cell_d2 = get_writable_cell(ws, 2, 4)
    assert cell_d2.coordinate == "B1"
    cell_d2.value = "Updated"
    assert ws["B1"].value == "Updated"


def test_eod_duplicate_pos_sum_and_latest_toggle():
    import tempfile
    import openpyxl
    from core.pos_recon.eod_extractor import extract_eod_source

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Daily_POS_2026_10_20"

    headers = ["Timestamp", "Submission ID", "Station", "Report Date", "Full POS ID", "Amount"]
    ws.append(headers)
    # POS 43 submission 1: 300 at 09:00:00
    ws.append(["2026-10-20 09:00:00", "SUB1", "Station Alpha", "2026-10-20", "SGR 43", 300])
    # POS 43 submission 2: 400 at 17:00:00 (latest)
    ws.append(["2026-10-20 17:00:00", "SUB2", "Station Alpha", "2026-10-20", "SGR 43", 400])
    # Distinct POS 10: 500 (single)
    ws.append(["2026-10-20 18:00:00", "SUB3", "Station Beta", "2026-10-20", "OBX 10", 500])

    # Also test out-of-order duplicate POS 88: row with later timestamp appears FIRST in sheet
    # Row 5 (TA 88): 850 at 16:30:00 (later timestamp)
    ws.append(["2026-10-20 16:30:00", "SUB4", "Station Gamma", "2026-10-20", "TA 88", 850])
    # Row 6 (TA 88): 250 at 08:15:00 (earlier timestamp, even though it appears later in file)
    ws.append(["2026-10-20 08:15:00", "SUB5", "Station Gamma", "2026-10-20", "TA 88", 250])

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        wb.save(tmp.name)
        tmp_path = Path(tmp.name)

    try:
        # 1. Default mode (use_latest_duplicate=True):
        # - SGR 43: latest timestamp is 17:00:00 -> 400.0
        # - TA 88: latest timestamp is 16:30:00 -> 850.0 (despite row order)
        # - OBX 10: non-duplicate -> 500.0
        res_latest = extract_eod_source(tmp_path, "2026-10-20")
        assert res_latest["amountByPos"]["SGR 43"] == 400.0
        assert res_latest["amountByPos"]["TA 88"] == 850.0
        assert res_latest["amountByPos"]["OBX 10"] == 500.0
        assert len(res_latest["duplicateTerminals"]) == 2

        dup_43 = next(d for d in res_latest["duplicateTerminals"] if d["pos"] == "SGR 43")
        assert dup_43["count"] == 2
        assert dup_43["computedAmount"] == 400.0
        assert dup_43["latestTimestamp"] == "2026-10-20 17:00:00"
        assert dup_43["latestSubmissionId"] == "SUB2"
        assert dup_43["totalSum"] == 700.0

        dup_88 = next(d for d in res_latest["duplicateTerminals"] if d["pos"] == "TA 88")
        assert dup_88["count"] == 2
        assert dup_88["computedAmount"] == 850.0
        assert dup_88["latestTimestamp"] == "2026-10-20 16:30:00"
        assert dup_88["latestSubmissionId"] == "SUB4"
        assert dup_88["totalSum"] == 1100.0

        # 2. Disabled mode (use_latest_duplicate=False):
        # - SGR 43: 300 + 400 = 700.0
        # - TA 88: 850 + 250 = 1100.0
        res_sum = extract_eod_source(tmp_path, "2026-10-20", use_latest_duplicate=False)
        assert res_sum["amountByPos"]["SGR 43"] == 700.0
        assert res_sum["amountByPos"]["TA 88"] == 1100.0
        assert res_sum["amountByPos"]["OBX 10"] == 500.0

        # Backwards compatibility with sum_duplicates=True parameter
        res_compat = extract_eod_source(tmp_path, "2026-10-20", sum_duplicates=True)
        assert res_compat["amountByPos"]["SGR 43"] == 700.0
        assert res_compat["amountByPos"]["TA 88"] == 1100.0

    finally:
        cleanup_files(tmp_path)


def test_eod_real_file_duplicate_latest_resolution():
    from core.pos_recon.eod_extractor import extract_eod_source

    real_eod = Path("/home/im-redmond/Downloads/Moniepoint EOD 2026-10.xlsx")
    if not real_eod.exists():
        return

    # In Daily_POS_2026_10_13:
    # OBX 9:
    # Row 130: TS 2026-10-14T09:32:37+01:00, Amt=2004220.0
    # Row 157: TS 2026-10-14T09:42:42+01:00 (latest), Amt=2000000.0
    res = extract_eod_source(real_eod, "2026-10-13")
    assert "OBX 9" in res["amountByPos"]
    assert res["amountByPos"]["OBX 9"] == 2000000.0
    assert len(res["duplicateTerminals"]) == 1
    dup = res["duplicateTerminals"][0]
    assert dup["pos"] == "OBX 9"
    assert dup["computedAmount"] == 2000000.0
    assert dup["latestTimestamp"] == "2026-10-14T09:42:42+01:00"
    assert dup["latestSubmissionId"] == "SUB-ST042-20261013"


def test_eod_entry_mode_synthetic():
    import tempfile
    import openpyxl
    from core.pos_recon.eod_extractor import extract_eod_source

    wb = openpyxl.Workbook()
    
    # Sheet 1: Daily_POS_2026_10_18 (Yesterday)
    ws18 = wb.active
    ws18.title = "Daily_POS_2026_10_18"
    headers = ["Timestamp", "Submission ID", "Station", "Report Date", "Full POS ID", "Amount", "Entry Mode"]
    ws18.append(headers)
    # Day 18 DEFAULT (should be ignored when extracting Day 19)
    ws18.append(["2026-10-19 08:00:00", "SUB-D18-1", "Station 1", "2026-10-18", "SGR 99", 100, "DEFAULT"])
    # Day 18 CURRENT (eligible for Day 19 fill-in)
    # POS 10: missing on Day 19 -> will be adopted as fill-in
    ws18.append(["2026-10-18 12:00:00", "SUB-C18-1", "Station 1", "2026-10-19", "OBX 10", 500, "CURRENT"])
    # POS 20: duplicate with Day 19 DEFAULT (submitted yesterday as CURRENT for 800, today as DEFAULT for 950)
    ws18.append(["2026-10-18 13:00:00", "SUB-C18-2", "Station 2", "2026-10-19", "OBX 20", 800, "CURRENT"])
    # POS 30: missing on Day 19 -> will be adopted as fill-in
    ws18.append(["2026-10-18 14:00:00", "SUB-C18-3", "Station 3", "2026-10-19", "TA 30", 250, "CURRENT"])

    # Sheet 2: Daily_POS_2026_10_19 (Today)
    ws19 = wb.create_sheet("Daily_POS_2026_10_19")
    ws19.append(headers)
    # Day 19 DEFAULT
    # POS 20: today's DEFAULT amount is 950 -> MUST BE KEPT, 800 from yesterday discarded!
    ws19.append(["2026-10-20 09:00:00", "SUB-D19-1", "Station 2", "2026-10-19", "OBX 20", 950, "DEFAULT"])
    # POS 40: normal today DEFAULT
    ws19.append(["2026-10-20 09:30:00", "SUB-D19-2", "Station 4", "2026-10-19", "SGR 40", 1200, "DEFAULT"])
    # Day 19 CURRENT (is for Day 20, so must NOT be in Day 19)
    ws19.append(["2026-10-19 15:00:00", "SUB-C19-1", "Station 5", "2026-10-20", "SGR 50", 300, "CURRENT"])

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        wb.save(tmp.name)
        tmp_path = Path(tmp.name)

    try:
        res = extract_eod_source(tmp_path, "2026-10-19")
        assert res["entryModeEnabled"] is True
        assert res["defaultCount"] == 2  # OBX 20 and SGR 40
        assert res["currentFillinCount"] == 2  # OBX 10 and TA 30
        assert res["crossDayDuplicatesCount"] == 1  # OBX 20
        assert len(res["amountByPos"]) == 4

        # Verify exact amounts
        assert res["amountByPos"]["OBX 20"] == 950.0  # DEFAULT chosen!
        assert res["amountByPos"]["SGR 40"] == 1200.0  # DEFAULT chosen!
        assert res["amountByPos"]["OBX 10"] == 500.0  # Fill-in from yesterday's CURRENT
        assert res["amountByPos"]["TA 30"] == 250.0  # Fill-in from yesterday's CURRENT
        assert "SGR 50" not in res["amountByPos"]  # Day 19 CURRENT is for Day 20, correctly excluded
        assert "SGR 99" not in res["amountByPos"]  # Day 18 DEFAULT is for Day 18, correctly excluded

        # Cross-day duplicate details
        dup = res["crossDayDuplicates"][0]
        assert dup["pos"] == "OBX 20"
        assert dup["defaultAmount"] == 950.0
        assert dup["currentAmount"] == 800.0
        assert dup["chosenAmount"] == 950.0

    finally:
        cleanup_files(tmp_path)


def test_eod_entry_mode_real_file_day27():
    from core.pos_recon.eod_extractor import extract_eod_source

    real_eod = Path("/home/im-redmond/Downloads/Moniepoint EOD 2026-10 (2).xlsx")
    if not real_eod.exists():
        real_eod = Path("/home/im-redmond/Downloads/Moniepoint EOD 2026-10 (1).xlsx")
    if not real_eod.exists():
        return

    res = extract_eod_source(real_eod, "2026-10-27")
    assert res["entryModeEnabled"] is True
    assert res["defaultCount"] == 102
    assert res["currentFillinCount"] == 63
    assert res["crossDayDuplicatesCount"] == 13
    assert len(res["amountByPos"]) == 165

    # Check that cross-day duplicate OBX 18 uses DEFAULT (172,000) not CURRENT (5,000)
    assert res["amountByPos"]["OBX 18"] == 172000.0

    # Check fill-in OBX 154 adopted from Day 26 CURRENT
    assert res["amountByPos"]["OBX 154"] == 123300.0

    # Check AGS 10 -> SGR 10 remapping (default True)
    assert "SGR 10" in res["amountByPos"]
    assert res["amountByPos"]["SGR 10"] == 1032745.0
    assert "AGS 10" not in res["amountByPos"]

    # When remap_ags10 is False, it remains AGS 10
    res_no_remap = extract_eod_source(real_eod, "2026-10-27", remap_ags10=False)
    assert "AGS 10" in res_no_remap["amountByPos"]
    assert res_no_remap["amountByPos"]["AGS 10"] == 1032745.0
    assert "SGR 10" not in res_no_remap["amountByPos"]


def test_multi_bank_loader_synthetic():
    import tempfile
    import openpyxl
    from core.pos_recon.bank_loader import load_multiple_json_to_workbook

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Block 1 headers in row 1 & row 4
    ws.cell(1, 3, "2026-10-08")
    ws.cell(4, 1, "TERMINAL SERIALS")
    ws.cell(4, 2, "CASHBOOK EOD")
    ws.cell(4, 3, "BANK STATEMENT")
    ws.cell(4, 4, "DIFFERENCE")

    # Block 2 headers in row 1 & row 4
    ws.cell(1, 10, "2026-10-09")
    ws.cell(4, 8, "TERMINAL SERIALS")
    ws.cell(4, 9, "CASHBOOK EOD")
    ws.cell(4, 10, "BANK STATEMENT")
    ws.cell(4, 11, "DIFFERENCE")

    # Terminals in Block 1
    ws.cell(5, 1, "SGR 10")
    ws.cell(5, 2, 500)
    ws.cell(6, 1, "OBX 20")
    ws.cell(6, 2, 800)

    # Terminals in Block 2
    ws.cell(5, 8, "SGR 10")
    ws.cell(5, 9, 600)
    ws.cell(6, 8, "OBX 20")
    ws.cell(6, 9, 850)

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        wb.save(tmp.name)
        wb_path = Path(tmp.name)

    day8_json = [
        {"pos": "SGR 10", "status": "OK", "credit": "500.00"},
        {"pos": "OBX 20", "status": "FAILED", "error": "Timeout", "credit": ""},
    ]
    day9_json = [
        {"pos": "SGR 10", "status": "OK", "credit": "600.00"},
        {"pos": "OBX 20", "status": "OK", "credit": "850.00"},
    ]

    items = [
        {"json_input": day8_json, "date": "2026-10-08", "block_number": 1, "filename": "day8.json"},
        {"json_input": day9_json, "date": "2026-10-09", "block_number": 2, "filename": "day9.json"},
    ]

    try:
        out_path, stats = load_multiple_json_to_workbook(wb_path, items)
        assert out_path.exists()
        assert stats["totalFiles"] == 2
        assert stats["totalMatched"] == 4
        assert stats["totalOkWritten"] == 3
        assert stats["totalStatusWritten"] == 1

        # Check values in output workbook
        check_wb = openpyxl.load_workbook(out_path)
        check_ws = check_wb.active
        # Block 1: SGR 10 -> 500, OBX 20 -> FAILED
        assert check_ws.cell(5, 3).value == 500.0
        assert check_ws.cell(6, 3).value == "FAILED"
        assert check_ws.cell(6, 3).comment is not None

        # Block 2: SGR 10 -> 600, OBX 20 -> 850
        assert check_ws.cell(5, 10).value == 600.0
        assert check_ws.cell(6, 10).value == 850.0

        check_wb.close()
        cleanup_files(out_path)
    finally:
        cleanup_files(wb_path)






