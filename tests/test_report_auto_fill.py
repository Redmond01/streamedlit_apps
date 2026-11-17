"""Unit tests for Report Auto-Fill Engine."""

from datetime import datetime
import unittest
import openpyxl

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
        """Test batch auto-fill and zip archive generation."""
        files = [SAMPLE_MOWE, SAMPLE_ELEKO]
        results, zip_buf, summaries = auto_fill_multiple_station_files(
            files,
            target_sheet_mode="latest",
        )

        self.assertEqual(len(results), 2)
        self.assertEqual(len(summaries), 2)
        self.assertEqual(summaries[0].status, "Success")
        self.assertEqual(summaries[1].status, "Success")

        import zipfile
        with zipfile.ZipFile(zip_buf, "r") as zf:
            namelist = zf.namelist()
            self.assertEqual(len(namelist), 2)
            self.assertTrue(any("filled" in n for n in namelist))


if __name__ == "__main__":
    unittest.main()
