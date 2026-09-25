from __future__ import annotations

import logging
from pathlib import Path
import re
import uuid
import zipfile

logger = logging.getLogger(__name__)


def preserve_gear_extensions(
    source_master_path: Path,
    generated_output_path: Path,
    gear_sheet_name: str = "Sheet1",
    gear_percent_row: int | None = None,
    average_gear_row: int | None = None,
) -> bool:
    """Restores modern Excel OpenXML extensions (such as x14:conditionalFormattings) stripped by openpyxl.

    Openpyxl drops the <extLst> block containing enhanced conditional formatting rules
    (UserWarning: Conditional Formatting extension is not supported and will be removed).
    This function extracts the original <extLst> from the template and re-injects it into
    the generated output workbook, ensuring all color highlights (green, yellow, red) are preserved.
    """
    if not source_master_path.exists() or not generated_output_path.exists():
        return False

    try:
        with zipfile.ZipFile(source_master_path, "r") as z_orig:
            namelist_orig = z_orig.namelist()
            sheet_target = "xl/worksheets/sheet1.xml"
            if sheet_target not in namelist_orig:
                for name in namelist_orig:
                    if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
                        sheet_target = name
                        break
            if sheet_target not in namelist_orig:
                return False

            orig_sheet = z_orig.read(sheet_target).decode("utf-8", errors="ignore")
            orig_styles = (
                z_orig.read("xl/styles.xml").decode("utf-8", errors="ignore")
                if "xl/styles.xml" in namelist_orig
                else ""
            )

        extLst_match = re.search(r"<extLst>.*?</extLst>", orig_sheet, re.DOTALL)
        if not extLst_match:
            return False

        extLst = extLst_match.group(0)

        # If new target rows need conditional formatting rules, ensure they exist in extLst
        if gear_percent_row is not None or average_gear_row is not None:
            extLst = _ensure_rows_in_extlst(extLst, gear_percent_row, average_gear_row)

        # Read generated output zip
        with zipfile.ZipFile(generated_output_path, "r") as z_gen:
            files = {name: z_gen.read(name) for name in z_gen.namelist()}

        if sheet_target not in files:
            return False

        sheet_xml = files[sheet_target].decode("utf-8", errors="ignore")

        # Add required namespaces if missing from <worksheet>
        namespaces = (
            ' xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"'
            ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
            ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"'
            ' xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"'
            ' mc:Ignorable="x14ac"'
        )
        if "xmlns:x14" not in sheet_xml:
            sheet_xml = re.sub(r"(<worksheet\b[^>]*)>", r"\1" + namespaces + ">", sheet_xml, count=1)

        # Inject extLst right before </worksheet>
        if "</worksheet>" in sheet_xml and "<extLst>" not in sheet_xml:
            sheet_xml = sheet_xml.replace("</worksheet>", extLst + "</worksheet>")

        files[sheet_target] = sheet_xml.encode("utf-8")

        # Also preserve extLst in xl/styles.xml if present in original
        styles_ext_match = re.search(r"<extLst>.*?</extLst>", orig_styles, re.DOTALL)
        if styles_ext_match and "xl/styles.xml" in files:
            styles_xml = files["xl/styles.xml"].decode("utf-8", errors="ignore")
            if "</styleSheet>" in styles_xml and "<extLst>" not in styles_xml:
                styles_xml = styles_xml.replace("</styleSheet>", styles_ext_match.group(0) + "</styleSheet>")
                files["xl/styles.xml"] = styles_xml.encode("utf-8")

        # Write back into generated_output_path atomically
        temp_zip = generated_output_path.with_name(f"{generated_output_path.name}.tmp")
        with zipfile.ZipFile(temp_zip, "w", compression=zipfile.ZIP_DEFLATED) as z_final:
            for name, data in files.items():
                z_final.writestr(name, data)

        temp_zip.replace(generated_output_path)
        logger.info("Successfully restored conditional formatting extensions to %s", generated_output_path)
        return True

    except Exception as exc:
        logger.warning("Failed to restore Excel extensions: %s", exc)
        return False


def _ensure_rows_in_extlst(extLst: str, gear_percent_row: int | None, average_gear_row: int | None) -> str:
    """Clones conditional formatting rules for newly targeted day rows if not already present."""
    new_rules: list[str] = []

    # Find an existing rule pattern for gear percent and average gear
    percent_match = re.findall(r"<x14:conditionalFormatting\b.*?<xm:sqref>[A-Z0-9:]*134[A-Z0-9:]*</xm:sqref></x14:conditionalFormatting>", extLst, re.DOTALL)
    average_match = re.findall(r"<x14:conditionalFormatting\b.*?<xm:sqref>[A-Z0-9:]*135[A-Z0-9:]*</xm:sqref></x14:conditionalFormatting>", extLst, re.DOTALL)

    if gear_percent_row is not None and f":IW{gear_percent_row}</xm:sqref>" not in extLst and percent_match:
        for block in percent_match:
            cloned = re.sub(r'id="\{[^\}]+\}"', lambda m: f'id="{{{str(uuid.uuid4()).upper()}}}"', block)
            cloned = re.sub(r"<xm:sqref>.*?</xm:sqref>", f"<xm:sqref>A{gear_percent_row}:IW{gear_percent_row}</xm:sqref>", cloned)
            new_rules.append(cloned)

    if average_gear_row is not None and f":IW{average_gear_row}</xm:sqref>" not in extLst and average_match:
        for block in average_match:
            cloned = re.sub(r'id="\{[^\}]+\}"', lambda m: f'id="{{{str(uuid.uuid4()).upper()}}}"', block)
            cloned = re.sub(r"<xm:sqref>.*?</xm:sqref>", f"<xm:sqref>A{average_gear_row}:IW{average_gear_row}</xm:sqref>", cloned)
            new_rules.append(cloned)

    if new_rules:
        insert_marker = "</x14:conditionalFormattings>"
        if insert_marker in extLst:
            extLst = extLst.replace(insert_marker, "".join(new_rules) + insert_marker)

    return extLst
