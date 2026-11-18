"""Unit tests for Report Auto-Fill Engine."""

from datetime import datetime
from pathlib import Path
import unittest
import openpyxl

from core.report_sync.report_syncer import clean_report_name
from core.report_sync.report_auto_fill import (
    auto_fill_multiple_station_files,
    auto_fill_sheet,
    auto_fill_station_file,
    auto_fill_workbook,
)

SAMPLE_MOWE = "/home/im-redmond/.gemini/antigravity/brain/082953db-c1e9-4812-9851-757b0067cd34/.user_uploaded/media_1794918520855_94175480.xlsx"
SAMPLE_ELEKO = "/home/im-redmond/.gemini/antigravity/brain/082953db-c1e9-4812-9851-757b0067cd34/.user_uploaded/media_1794918520850_6408ad0b.xlsx"


class TestReportAutoFill(unittest.TestCase):
    def test_auto_fill_mowe_sheet2(self):
        """Test outer to inner copy on Mowe Sheet 2 (tanks, products received, pos)."""
        wb = openpyxl.load_workbook(SAMPLE_MOWE, data_only=False)
        ws = wb["Sheet2"]

        # Before fill: check outer product received in U10 and U11
        self.assertEqual(ws["U10"].value, 21100)
        self.assertEqual(ws["U11"].value, 29650)
        outer_pms1_open = ws["E20"].value
        outer_pms1_close = ws["F20"].value

        counts = auto_fill_sheet(ws)

        self.assertEqual(counts["tanks"], 5)
        self.assertGreaterEqual(counts["products"], 2)
        self.assertGreater(counts["pos"], 0)

        # 1. Tank Dipping copied to inner table (Cols AH, AI)
        self.assertEqual(ws["AH20"].value, outer_pms1_open)
        self.assertEqual(ws["AI20"].value, outer_pms1_close)

        # 2. Product Received copied to inner table (Cols AW, AX)
        self.assertEqual(ws["AX10"].value, 21100)
        self.assertEqual(ws["AX11"].value, 29650)

        # 3. Matching POS copied to inner table (Col AE is pos3)
        # Left side had F27 = 727275 and F29 = 3059059 -> combined = 3786334
        self.assertEqual(ws["AE27"].value, 3786334)

    def test_auto_fill_eleko_sheet1(self):
        """Test outer to inner copy on Eleko 2 Sheet 1 (tanks, matching POS terminals)."""
        wb = openpyxl.load_workbook(SAMPLE_ELEKO, data_only=False)
        ws = wb["Sheet1"]

        outer_pms1_open = ws["E17"].value
        outer_pms1_close = ws["F17"].value

        counts = auto_fill_sheet(ws)

        self.assertEqual(counts["tanks"], 3)
        self.assertGreater(counts["pos"], 0)

        # Tank dipping copied to inner (Cols AE, AF)
        self.assertEqual(ws["AE17"].value, outer_pms1_open)
        self.assertEqual(ws["AF17"].value, outer_pms1_close)

        # Matching POS copied: Left had G22 (SGR 358) = 1674870.86
        # Right SGR 358 is at AF22
        self.assertEqual(ws["AF22"].value, 1674870.86)

    def test_auto_fill_batch_and_zip(self):
        """Test batch auto-fill and zip archive generation with clean report names."""
        files = [SAMPLE_MOWE, SAMPLE_ELEKO]
        results, zip_buf, summaries = auto_fill_multiple_station_files(
            files,
            target_sheet_mode="latest",
        )

        self.assertEqual(len(results), 2)
        self.assertEqual(len(summaries), 2)
        self.assertEqual(summaries[0].status, "Success")
        self.assertEqual(summaries[1].status, "Success")

        # Verify output filename format is clean as report name (no _filled, no _sync, no timestamps)
        expected_name_0 = f"{clean_report_name(SAMPLE_MOWE)}.xlsx"
        expected_name_1 = f"{clean_report_name(SAMPLE_ELEKO)}.xlsx"
        self.assertEqual(results[0][0], expected_name_0)
        self.assertEqual(results[1][0], expected_name_1)

        import zipfile
        with zipfile.ZipFile(zip_buf, "r") as zf:
            namelist = zf.namelist()
            self.assertEqual(len(namelist), 2)
            self.assertIn(expected_name_0, namelist)
            self.assertIn(expected_name_1, namelist)
            for n in namelist:
                self.assertNotIn("filled", n)
                self.assertNotIn("sync", n)

    def test_auto_fill_cleans_synced_input_filename(self):
        """Test that an input named with '_sync' or '_synced_20261117' produces clean report name."""
        import io

        with open(SAMPLE_MOWE, "rb") as f:
            data = f.read()

        mock_f = io.BytesIO(data)
        mock_f.name = "NOVEMBER_Mowe_sgr_sync.xlsx"
        out_bytes, out_name, summaries = auto_fill_station_file(mock_f)

        self.assertEqual(out_name, "NOVEMBER_Mowe_sgr.xlsx")
        self.assertEqual(summaries[0].station_name, "NOVEMBER_Mowe_sgr")

    def test_auto_fill_fallbacks_oyo_ilora_tank_row32(self):
        """Test outer-to-inner copy on Oyo ilora where tank header is on row 32."""
        p = Path("/home/im-redmond/.gemini/antigravity/brain/082953db-c1e9-4812-9851-757b0067cd34/scratch/all_stations_synced/NOVEMBER Oyo ilora obx_synced_20261117.xlsx")
        if not p.exists():
            return
        _, out_name, summaries = auto_fill_station_file(p, target_sheet_mode="latest")
        self.assertEqual(summaries[0].tanks_copied, 10)
        self.assertEqual(summaries[0].pos_copied, 16)

    def test_auto_fill_fallbacks_ogere1_addup(self):
        """Test outer-to-inner copy on Ogere 1 where remittance header is ADDUP."""
        p = Path("/home/im-redmond/.gemini/antigravity/brain/082953db-c1e9-4812-9851-757b0067cd34/scratch/all_stations_synced/NOVEMBER Ogere 1 sgr_synced_20261117.xlsx")
        if not p.exists():
            return
        _, out_name, summaries = auto_fill_station_file(p, target_sheet_mode="latest")
        self.assertEqual(summaries[0].tanks_copied, 4)
        self.assertEqual(summaries[0].pos_copied, 17)

    def test_auto_fill_single_sided_jah_michael(self):
        """Test that single-sided stations without an inner meter side are gracefully handled."""
        p = Path("/home/im-redmond/.gemini/antigravity/brain/082953db-c1e9-4812-9851-757b0067cd34/scratch/all_stations_synced/NOVEMBER Jah michael Ta_synced_20261117.xlsx")
        if not p.exists():
            return
        _, out_name, summaries = auto_fill_station_file(p, target_sheet_mode="latest")
        self.assertEqual(summaries[0].status, "Success")
        self.assertIn("Single-sided", summaries[0].message)

    def test_auto_fill_dynamic_tanks_guru(self):
        """Test outer-to-inner auto-fill on dynamic multi-pair station Guru."""
        p = Path("/home/im-redmond/Downloads/Station_Reports_Synced_20261118_144442/NOVEMBER Guru obx_sync.xlsx")
        if not p.exists():
            return
        wb = openpyxl.load_workbook(p, data_only=False)
        ws = wb["Sheet16"]
        counts = auto_fill_sheet(ws)
        self.assertGreaterEqual(counts["tanks"], 4)
        # PMS1 (row 19) in outer F19/G19 copied to inner AE19/AF19
        self.assertEqual(ws["AE19"].value, ws["F19"].value)
        self.assertEqual(ws["AF19"].value, ws["G19"].value)
        # Manifolded formula in AE21/AF21 preserved intact
        self.assertEqual(ws["AE21"].value, "=AC20+AC21")
        self.assertEqual(ws["AF21"].value, "=AD20+AD21")

    def test_auto_fill_dynamic_tanks_epe2(self):
        """Test outer-to-inner auto-fill and formula preservation on Epe 2."""
        p = Path("/home/im-redmond/Downloads/Station_Reports_Synced_20261118_144442/NOVEMBER Epe 2 sgr_sync.xlsx")
        if not p.exists():
            return
        wb = openpyxl.load_workbook(p, data_only=False)
        ws = wb["Sheet16"]
        counts = auto_fill_sheet(ws)
        self.assertGreaterEqual(counts["tanks"], 5)
        # PMS2 formula in AG19/AH19 preserved intact
        self.assertEqual(ws["AG19"].value, "=AE19+AE18")
        self.assertEqual(ws["AH19"].value, "=AF19+AF18")

    def test_auto_fill_meters_healed_and_filled_lekki1(self):
        """Test that zero-meters on Lekki 1 are healed from previous sheets and inner closing is filled."""
        p = Path("/home/im-redmond/Downloads/Station_Reports_Synced_20261118_144442/NOVEMBER Lekki 1 sgr_sync.xlsx")
        if not p.exists():
            return
        _, out_name, summaries = auto_fill_station_file(p, target_sheet_mode="latest")
        self.assertEqual(out_name, "NOVEMBER Lekki 1 sgr.xlsx")
        self.assertEqual(summaries[0].status, "Success")
        self.assertGreaterEqual(summaries[0].pumps_synced, 1)


if __name__ == "__main__":
    unittest.main()
