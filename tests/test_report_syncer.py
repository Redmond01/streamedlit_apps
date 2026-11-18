"""Unit tests for Report Auto-Sync Engine."""

from datetime import datetime
import io
import unittest
from pathlib import Path
import openpyxl

from core.report_sync.report_syncer import (
    clean_report_name,
    sync_multiple_station_files,
    sync_station_file,
    sync_station_report,
)

SAMPLE_MOWE = "/home/im-redmond/.gemini/antigravity/brain/082953db-c1e9-4812-9851-757b0067cd34/.user_uploaded/media_1794918520855_94175480.xlsx"
SAMPLE_ELEKO = "/home/im-redmond/.gemini/antigravity/brain/082953db-c1e9-4812-9851-757b0067cd34/.user_uploaded/media_1794918520850_6408ad0b.xlsx"
SAMPLE_AGBARA = "/home/im-redmond/Downloads/AUGUST Agbara 1  sgr (2).xlsx"
SAMPLE_02UI_SINGLE = "/home/im-redmond/.gemini/antigravity/brain/082953db-c1e9-4812-9851-757b0067cd34/.user_uploaded/media_1795001747826_fad217c7.xlsx"
SAMPLE_02UI_DOUBLE = "/home/im-redmond/.gemini/antigravity/brain/082953db-c1e9-4812-9851-757b0067cd34/.user_uploaded/media_1795003760359_52f9ad8c.xlsx"


class TestReportSyncer(unittest.TestCase):
    def test_sync_mowe_report(self):
        """Test rollover of Mowe station workbook (10 pumps, 5 tanks)."""
        wb = openpyxl.load_workbook(SAMPLE_MOWE, data_only=False)
        wb_data = openpyxl.load_workbook(SAMPLE_MOWE, data_only=True)
        old_sheet_count = len(wb.worksheets)
        src_ws = wb.worksheets[-1]
        self.assertEqual(src_ws.title, "Sheet10")

        # Capture closing values from Sheet10
        pms1_close = src_ws["F3"].value
        inner_pms1_close = src_ws["AI3"].value
        pms1_tank_close = src_ws["F20"].value

        updated_wb, summary = sync_station_report(wb, wb_data=wb_data)

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
        # ADD UP brought forward as evaluated float of previous day's AVAILABLE CASH (never an equation)
        self.assertEqual(new_ws["Q27"].value, 58868374.93)
        self.assertFalse(str(new_ws["Q27"].value).startswith("="))
        # AVAILABLE CASH keeps its formula
        self.assertTrue(str(new_ws["R27"].value).startswith("="))
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
        wb_data = openpyxl.load_workbook(SAMPLE_ELEKO, data_only=True)
        src_ws = wb.worksheets[-1]
        self.assertEqual(src_ws.title, "Sheet10")

        pms1_close = src_ws["E3"].value
        tank_close = src_ws["F17"].value

        updated_wb, summary = sync_station_report(wb, wb_data=wb_data)

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
        # ADD UP brought forward as evaluated float of previous day's AVAILABLE CASH (never an equation)
        self.assertEqual(new_ws["O22"].value, 19092865.7)
        self.assertFalse(str(new_ws["O22"].value).startswith("="))
        # AVAILABLE CASH keeps its formula
        self.assertTrue(str(new_ws["P22"].value).startswith("="))
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

    def test_sync_02ui_single_sided_report(self):
        """Test rollover of 02UI single-sided report (Sheet16 -> Sheet17, OPEN DIPPG detection and Row 11 AGO 2 reset)."""
        if not Path(SAMPLE_02UI_SINGLE).exists():
            return
        wb = openpyxl.load_workbook(SAMPLE_02UI_SINGLE, data_only=False)
        wb_d = openpyxl.load_workbook(SAMPLE_02UI_SINGLE, data_only=True)
        if "Sheet17" in wb.sheetnames:
            wb.remove(wb["Sheet17"])
        if "Sheet17" in wb_d.sheetnames:
            wb_d.remove(wb_d["Sheet17"])

        updated_wb, summary = sync_station_report(wb, wb_data=wb_d)
        self.assertEqual(summary.outer_pumps_updated, 7)
        self.assertEqual(summary.tanks_updated, 3)

        new_ws = updated_wb.worksheets[-1]
        self.assertEqual(new_ws.title, "Sheet17")

        # Row 11 AGO 2 must be reset
        self.assertEqual(new_ws["D11"].value, 776872.03)
        self.assertEqual(new_ws["E11"].value, 776872.03)

        # Tank section (Rows 16-18): Formulas preserved and dippings reset
        self.assertEqual(new_ws["D16"].value, 1400)
        self.assertEqual(new_ws["E16"].value, 1400)
        self.assertEqual(new_ws["F16"].value, "=(D16+S7)-E16")

        self.assertEqual(new_ws["D17"].value, 17100)
        self.assertEqual(new_ws["E17"].value, 17100)
        self.assertEqual(new_ws["F17"].value, "=(D17+S8)-E17")

        self.assertEqual(new_ws["D18"].value, 13250)
        self.assertEqual(new_ws["E18"].value, 13250)
        self.assertEqual(new_ws["F18"].value, "=(D18+S9)-E18")

    def test_sync_02ui_double_sided_report(self):
        """Test rollover of 02UI double-sided report with Row 11 AGO 2 outer & inner meters and 5 tanks."""
        if not Path(SAMPLE_02UI_DOUBLE).exists():
            return
        wb = openpyxl.load_workbook(SAMPLE_02UI_DOUBLE, data_only=False)
        wb_d = openpyxl.load_workbook(SAMPLE_02UI_DOUBLE, data_only=True)
        if "Sheet17" in wb.sheetnames:
            wb.remove(wb["Sheet17"])
        if "Sheet17" in wb_d.sheetnames:
            wb_d.remove(wb_d["Sheet17"])

        updated_wb, summary = sync_station_report(wb, wb_data=wb_d)
        self.assertEqual(summary.outer_pumps_updated, 7)
        self.assertEqual(summary.inner_pumps_updated, 7)
        self.assertEqual(summary.tanks_updated, 5)

        new_ws = updated_wb.worksheets[-1]
        self.assertEqual(new_ws.title, "Sheet17")

        # Row 11 outer and inner meter must both be reset to closing meter
        self.assertEqual(new_ws["F11"].value, 241731.9)
        self.assertEqual(new_ws["G11"].value, 241731.9)
        self.assertEqual(new_ws["H11"].value, 241731.9)

        self.assertEqual(new_ws["AF11"].value, 241832)
        self.assertEqual(new_ws["AG11"].value, 241832)
        self.assertEqual(new_ws["AH11"].value, 241832)

    def test_clean_report_name(self):
        """Test clean_report_name regex strips suffixes and timestamps correctly."""
        cases = {
            "NOVEMBER_Mowe_sgr.xlsx": "NOVEMBER_Mowe_sgr",
            "NOVEMBER_Mowe_sgr_sync.xlsx": "NOVEMBER_Mowe_sgr",
            "NOVEMBER_Mowe_sgr_synced.xlsx": "NOVEMBER_Mowe_sgr",
            "NOVEMBER_Mowe_sgr_synced_20261117.xlsx": "NOVEMBER_Mowe_sgr",
            "NOVEMBER_Mowe_sgr_sync_20261117.xlsx": "NOVEMBER_Mowe_sgr",
            "NOVEMBER_Mowe_sgr_filled.xlsx": "NOVEMBER_Mowe_sgr",
            "NOVEMBER_Mowe_sgr_filled_20261117_143640.xlsx": "NOVEMBER_Mowe_sgr",
            "NOVEMBER_Mowe_sgr_synced_20261111_filled_20261117_143640.xlsx": "NOVEMBER_Mowe_sgr",
            "NOVEMBER_Mowe_sgr_reset.xlsx": "NOVEMBER_Mowe_sgr",
            "NOVEMBER Jah michael Ta_synced_20261117.xlsx": "NOVEMBER Jah michael Ta",
            "NOVEMBER_Iwo_2_obx_synced_20261117.xlsx": "NOVEMBER_Iwo_2_obx",
        }
        for raw, expected in cases.items():
            self.assertEqual(clean_report_name(raw), expected)

    def test_sync_multiple_station_files_and_zip(self):
        """Test batch processing and zip packaging with '_sync.xlsx' output naming."""
        files = [SAMPLE_MOWE, SAMPLE_ELEKO]
        results, zip_buf, summaries = sync_multiple_station_files(files)

        self.assertEqual(len(results), 2)
        self.assertEqual(len(summaries), 2)
        self.assertEqual(summaries[0].status, "Success")
        self.assertEqual(summaries[1].status, "Success")

        # Verify output filename format contains "_sync.xlsx"
        expected_name_0 = f"{clean_report_name(SAMPLE_MOWE)}_sync.xlsx"
        expected_name_1 = f"{clean_report_name(SAMPLE_ELEKO)}_sync.xlsx"
        self.assertEqual(results[0][0], expected_name_0)
        self.assertEqual(results[1][0], expected_name_1)

        # Verify zip archive
        import zipfile
        with zipfile.ZipFile(zip_buf, "r") as zf:
            namelist = zf.namelist()
            self.assertEqual(len(namelist), 2)
            self.assertIn(expected_name_0, namelist)
            self.assertIn(expected_name_1, namelist)

    def test_sync_fallbacks_ogbomoso2_meters(self):
        """Test meter detection fallback when outer/inner opening headers are labeled CLOSING METER."""
        p = Path("/home/im-redmond/.gemini/antigravity/brain/082953db-c1e9-4812-9851-757b0067cd34/scratch/all_stations_synced/NOVEMBER Ogbomoso 2 obx_synced_20261117.xlsx")
        if not p.exists():
            return
        wb = openpyxl.load_workbook(p, data_only=False)
        wb_d = openpyxl.load_workbook(p, data_only=True)
        if len(wb.worksheets) > 1:
            wb.remove(wb.worksheets[-1])
            wb_d.remove(wb_d.worksheets[-1])

        updated_wb, summary = sync_station_report(wb, wb_data=wb_d)
        self.assertEqual(summary.outer_pumps_updated, 5)
        self.assertEqual(summary.inner_pumps_updated, 5)
        self.assertEqual(summary.tanks_updated, 4)

        new_ws = updated_wb.worksheets[-1]
        # Check that E3 (open) was reset to F3 (close)
        self.assertEqual(new_ws["E3"].value, new_ws["F3"].value)
        self.assertEqual(new_ws["AE3"].value, new_ws["AF3"].value)

    def test_sync_fallbacks_oyo_ilora_tank_row32(self):
        """Test tank detection fallback when tank headers start on row 32 (large pump count)."""
        p = Path("/home/im-redmond/.gemini/antigravity/brain/082953db-c1e9-4812-9851-757b0067cd34/scratch/all_stations_synced/NOVEMBER Oyo ilora obx_synced_20261117.xlsx")
        if not p.exists():
            return
        wb = openpyxl.load_workbook(p, data_only=False)
        wb_d = openpyxl.load_workbook(p, data_only=True)
        if len(wb.worksheets) > 1:
            wb.remove(wb.worksheets[-1])
            wb_d.remove(wb_d.worksheets[-1])

        updated_wb, summary = sync_station_report(wb, wb_data=wb_d)
        self.assertEqual(summary.outer_pumps_updated, 16)
        self.assertEqual(summary.inner_pumps_updated, 16)
        self.assertEqual(summary.tanks_updated, 10)

        new_ws = updated_wb.worksheets[-1]
        # Check tank row 34 opening dipping == closing dipping
        self.assertEqual(new_ws["F34"].value, new_ws["G34"].value)

    def test_sync_fallbacks_ogere1_addup(self):
        """Test remittance detection fallback when header is labeled ADDUP without space."""
        p = Path("/home/im-redmond/.gemini/antigravity/brain/082953db-c1e9-4812-9851-757b0067cd34/scratch/all_stations_synced/NOVEMBER Ogere 1 sgr_synced_20261117.xlsx")
        if not p.exists():
            return
        wb = openpyxl.load_workbook(p, data_only=False)
        wb_d = openpyxl.load_workbook(p, data_only=True)
        if len(wb.worksheets) > 1:
            wb.remove(wb.worksheets[-1])
            wb_d.remove(wb_d.worksheets[-1])

        updated_wb, summary = sync_station_report(wb, wb_data=wb_d)
        self.assertEqual(summary.outer_pumps_updated, 16)
        self.assertEqual(summary.inner_pumps_updated, 16)
        self.assertGreater(summary.pos_cleared, 0)

    def test_sync_fallbacks_interchange_date_row2(self):
        """Test date recognition fallback when date cell is in row 2 (B2)."""
        p = Path("/home/im-redmond/.gemini/antigravity/brain/082953db-c1e9-4812-9851-757b0067cd34/scratch/all_stations_synced/NOVEMBER Interchange sgr_synced_Unknown.xlsx")
        if not p.exists():
            return
        wb = openpyxl.load_workbook(p, data_only=False)
        wb_d = openpyxl.load_workbook(p, data_only=True)
        if len(wb.worksheets) > 1:
            wb.remove(wb.worksheets[-1])
            wb_d.remove(wb_d.worksheets[-1])

        updated_wb, summary = sync_station_report(wb, wb_data=wb_d)
        self.assertNotEqual(summary.new_date, "Unknown")
        self.assertEqual(summary.new_date, "2026-11-17")

    def test_sync_dynamic_tanks_guru(self):
        """Test multi-column dynamic tanks in Guru station (cols D/E, F/G, combination formulas)."""
        p = Path("/home/im-redmond/Downloads/Station_Reports_Synced_20261118_144442/NOVEMBER Guru obx_sync.xlsx")
        if not p.exists():
            return
        wb = openpyxl.load_workbook(p, data_only=False)
        wb_d = openpyxl.load_workbook(p, data_only=True)
        updated_wb, summary = sync_station_report(wb, wb_data=wb_d)
        self.assertEqual(summary.tanks_updated, 6)
        new_ws = updated_wb.worksheets[-1]
        # Standalone PMS1 in F19/G19 reset
        self.assertEqual(new_ws["F19"].value, new_ws["G19"].value)
        # Individual PMS2 in D20/E20 reset
        self.assertEqual(new_ws["D20"].value, new_ws["E20"].value)
        # Manifolded PMS3 formulas preserved in F21/G21
        self.assertEqual(new_ws["F21"].value, "=D20+D21")
        self.assertEqual(new_ws["G21"].value, "=E20+E21")
        # Standalone AGO1 in F24/G24 reset
        self.assertEqual(new_ws["F24"].value, new_ws["G24"].value)

    def test_sync_dynamic_tanks_epe2(self):
        """Test multi-column dynamic tanks in Epe 2 station."""
        p = Path("/home/im-redmond/Downloads/Station_Reports_Synced_20261118_144442/NOVEMBER Epe 2 sgr_sync.xlsx")
        if not p.exists():
            return
        wb = openpyxl.load_workbook(p, data_only=False)
        wb_d = openpyxl.load_workbook(p, data_only=True)
        updated_wb, summary = sync_station_report(wb, wb_data=wb_d)
        self.assertEqual(summary.tanks_updated, 6)
        new_ws = updated_wb.worksheets[-1]
        self.assertEqual(new_ws["D18"].value, new_ws["E18"].value)
        self.assertEqual(new_ws["F19"].value, "=D19+D18")
        self.assertEqual(new_ws["G19"].value, "=E19+E18")
        self.assertEqual(new_ws["F20"].value, new_ws["G20"].value)
        self.assertEqual(new_ws["F21"].value, new_ws["G21"].value)

    def test_sync_lekki1_meters_ago_lpg_pms(self):
        """Test that Lekki 1 inner meters for AGO/LPG and outer meters for PMS are correctly resolved and not zeroed out."""
        p = Path("/home/im-redmond/Downloads/Station_Reports_Synced_20261118_144442/NOVEMBER Lekki 1 sgr_sync.xlsx")
        if not p.exists():
            return
        wb = openpyxl.load_workbook(p, data_only=False)
        wb_d = openpyxl.load_workbook(p, data_only=True)
        updated_wb, summary = sync_station_report(wb, wb_data=wb_d)
        new_ws = updated_wb.worksheets[-1]

        # Inner meters for AGO 1 (Row 14), AGO 2 (Row 15), LPG PUMP (Row 19) must NOT be 0
        self.assertEqual(new_ws["AG14"].value, 304750)
        self.assertEqual(new_ws["AH14"].value, 304750)
        self.assertEqual(new_ws["AI14"].value, 304750)
        self.assertEqual(new_ws["AJ14"].value, 0)  # Inactive shift stays 0

        self.assertEqual(new_ws["AG15"].value, 89257)
        self.assertEqual(new_ws["AH15"].value, 89257)
        self.assertEqual(new_ws["AI15"].value, 89257)

        self.assertEqual(new_ws["AG19"].value, 9964)
        self.assertEqual(new_ws["AH19"].value, 9964)
        self.assertEqual(new_ws["AI19"].value, 9964)

        # Outer meters for PMS 3 (Row 6) and PMS 4 (Row 7) recovered from previous sheets
        self.assertEqual(new_ws["G6"].value, 225231.72)
        self.assertEqual(new_ws["H6"].value, 225231.72)
        self.assertEqual(new_ws["G7"].value, 209111.22)
        self.assertEqual(new_ws["H7"].value, 209111.22)


if __name__ == "__main__":
    unittest.main()
