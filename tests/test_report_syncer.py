"""Unit tests for Report Auto-Sync Engine."""

from datetime import datetime
import io
import unittest
from pathlib import Path
import openpyxl

from core.report_sync.report_syncer import (
    sync_multiple_station_files,
    sync_station_file,
    sync_station_report,
)

SAMPLE_MOWE = "/home/im-redmond/.gemini/antigravity/brain/082953db-c1e9-4812-9851-757b0067cd34/.user_uploaded/media_1794918520855_94175480.xlsx"
SAMPLE_ELEKO = "/home/im-redmond/.gemini/antigravity/brain/082953db-c1e9-4812-9851-757b0067cd34/.user_uploaded/media_1794918520850_6408ad0b.xlsx"
SAMPLE_AGBARA = "/home/im-redmond/Downloads/AUGUST Agbara 1  sgr (2).xlsx"


class TestReportSyncer(unittest.TestCase):
    def test_sync_mowe_report(self):
        """Test rollover of Mowe station workbook (10 pumps, 5 tanks)."""
        wb = openpyxl.load_workbook(SAMPLE_MOWE, data_only=False)
        old_sheet_count = len(wb.worksheets)
        src_ws = wb.worksheets[-1]
        self.assertEqual(src_ws.title, "Sheet10")

        # Capture closing values from Sheet10
        pms1_close = src_ws["F3"].value
        inner_pms1_close = src_ws["AI3"].value
        pms1_tank_close = src_ws["F20"].value
        avail_cash = src_ws["R27"].value

        updated_wb, summary = sync_station_report(wb)

        self.assertEqual(len(updated_wb.worksheets), old_sheet_count + 1)
        new_ws = updated_wb.worksheets[-1]
        self.assertEqual(new_ws.title, "Sheet11")
        self.assertEqual(summary.status, "Success")
        self.assertEqual(summary.outer_pumps_updated, 10)
        self.assertEqual(summary.inner_pumps_updated, 10)
        self.assertEqual(summary.tanks_updated, 5)

        # 1. Date: 2026-11-10 + 1 day = 2026-11-11
        date_val = new_ws["B1"].value
        if isinstance(date_val, datetime):
            self.assertEqual(date_val.strftime("%Y-%m-%d"), "2026-11-11")
        else:
            self.assertIn("2026-11-11", str(date_val))

        # 2. Outer Meter: Opening = Closing, Closing = Opening (Sales = 0)
        self.assertEqual(new_ws["E3"].value, pms1_close)
        self.assertEqual(new_ws["F3"].value, pms1_close)
        self.assertEqual(new_ws["J3"].value, "=F3-E3")

        # 3. Inner Meter: Opening = Closing, Closing = Opening
        self.assertEqual(new_ws["AH3"].value, inner_pms1_close)
        self.assertEqual(new_ws["AI3"].value, inner_pms1_close)

        # 4. Tank Dipping: Opening = Closing, Closing = Opening
        self.assertEqual(new_ws["E20"].value, pms1_tank_close)
        self.assertEqual(new_ws["F20"].value, pms1_tank_close)

        # 5. Products Received: Cleared
        self.assertIsNone(new_ws["U7"].value)
        self.assertIsNone(new_ws["AX7"].value)

        # 6. Expenses: Cleared
        self.assertIsNone(new_ws["Z8"].value)  # Imprest cleared

        # 7. POS: Values cleared
        self.assertIsNone(new_ws["F27"].value)
        self.assertIsNone(new_ws["G27"].value)
        # ADD UP brought forward statically from src_ws
        self.assertEqual(new_ws["Q27"].value, src_ws["Q27"].value)
        self.assertEqual(new_ws["AP27"].value, src_ws["AP27"].value)
        self.assertEqual(new_ws["AR27"].value, src_ws["AR27"].value)

        # Rollover Date updated to 2026-11-12
        rollover_date = new_ws["A29"].value
        if isinstance(rollover_date, datetime):
            self.assertEqual(rollover_date.strftime("%Y-%m-%d"), "2026-11-12")
        else:
            self.assertIn("2026-11-12", str(rollover_date))

    def test_sync_eleko_report(self):
        """Test rollover of Eleko 2 station workbook (7 pumps, 3 tanks)."""
        wb = openpyxl.load_workbook(SAMPLE_ELEKO, data_only=False)
        src_ws = wb.worksheets[-1]
        self.assertEqual(src_ws.title, "Sheet10")

        pms1_close = src_ws["E3"].value
        tank_close = src_ws["F17"].value

        updated_wb, summary = sync_station_report(wb)

        new_ws = updated_wb.worksheets[-1]
        self.assertEqual(new_ws.title, "Sheet11")
        self.assertEqual(summary.outer_pumps_updated, 7)
        self.assertEqual(summary.inner_pumps_updated, 7)
        self.assertEqual(summary.tanks_updated, 3)

        # Meter reset
        self.assertEqual(new_ws["D3"].value, pms1_close)
        self.assertEqual(new_ws["E3"].value, pms1_close)

        # Tank reset
        self.assertEqual(new_ws["E17"].value, tank_close)
        self.assertEqual(new_ws["F17"].value, tank_close)

        # POS cleared
        self.assertIsNone(new_ws["G22"].value)
        # ADD UP static
        self.assertEqual(new_ws["O22"].value, src_ws["O22"].value)
        self.assertEqual(new_ws["AO22"].value, src_ws["AO22"].value)

    def test_sync_agbara_report(self):
        """Test rollover of Agbara 1 station workbook (Sheet25 -> Sheet26)."""
        if not Path(SAMPLE_AGBARA).exists():
            return
        wb = openpyxl.load_workbook(SAMPLE_AGBARA, data_only=False)
        src_ws = wb.worksheets[-1]
        self.assertEqual(src_ws.title, "Sheet25")

        updated_wb, summary = sync_station_report(wb)
        new_ws = updated_wb.worksheets[-1]
        self.assertEqual(new_ws.title, "Sheet26")
        self.assertEqual(summary.outer_pumps_updated, 10)
        self.assertEqual(summary.tanks_updated, 4)

    def test_sync_multiple_station_files_and_zip(self):
        """Test batch processing and zip packaging."""
        files = [SAMPLE_MOWE, SAMPLE_ELEKO]
        results, zip_buf, summaries = sync_multiple_station_files(files)

        self.assertEqual(len(results), 2)
        self.assertEqual(len(summaries), 2)
        self.assertEqual(summaries[0].status, "Success")
        self.assertEqual(summaries[1].status, "Success")

        # Verify zip archive
        import zipfile
        with zipfile.ZipFile(zip_buf, "r") as zf:
            namelist = zf.namelist()
            self.assertEqual(len(namelist), 2)
            self.assertTrue(any("synced" in n for n in namelist))


if __name__ == "__main__":
    unittest.main()
