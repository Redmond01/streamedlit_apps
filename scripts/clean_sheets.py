"""Script and utility function to delete specific sheets (e.g. Sheet17, Sheet18)
from all station workbooks in a target folder.
"""

from pathlib import Path
import openpyxl


def delete_sheets_from_workbooks(
    directory: str | Path = "/home/im-redmond/Downloads/Station_Reports_Synced_20261118_144442",
    target_sheet_names: list[str] | tuple[str, ...] = ("Sheet17", "Sheet18"),
    delete_last_n_sheets: int = 0,
) -> dict[str, list[str]]:
    """Deletes target sheets (e.g. Sheet17 and Sheet18) from all Excel workbooks in a directory.

    Args:
        directory: Path to directory containing the .xlsx files.
        target_sheet_names: List or tuple of sheet names to remove (case-insensitive).
        delete_last_n_sheets: Optional; if > 0, also removes the last N sheets from each workbook.

    Returns:
        dict mapping filename to list of removed sheet names.
    """
    dir_path = Path(directory).expanduser().resolve()
    if not dir_path.exists():
        raise FileNotFoundError(f"Directory not found: {dir_path}")

    files = sorted([
        f for f in dir_path.glob("*.xlsx")
        if not f.name.startswith("~") and not f.name.startswith(".")
    ])

    normalized_targets = {name.strip().lower() for name in target_sheet_names}
    results: dict[str, list[str]] = {}

    print(f"Scanning {len(files)} workbooks in {dir_path}...\n")

    for idx, f in enumerate(files, 1):
        try:
            wb = openpyxl.load_workbook(f, data_only=False)
            removed: list[str] = []

            # 1. Match by explicit target sheet names (e.g. Sheet17, Sheet18)
            sheets_to_remove = [
                s for s in wb.sheetnames
                if s.strip().lower() in normalized_targets
            ]

            # 2. Check delete_last_n_sheets if specified
            if delete_last_n_sheets > 0:
                for s in wb.sheetnames[-delete_last_n_sheets:]:
                    if s not in sheets_to_remove:
                        sheets_to_remove.append(s)

            for s in sheets_to_remove:
                # Ensure at least 1 sheet remains in the workbook
                if len(wb.sheetnames) > 1 and s in wb.sheetnames:
                    wb.remove(wb[s])
                    removed.append(s)

            if removed:
                wb.save(f)
                results[f.name] = removed
                print(f"[{idx:02d}/{len(files)}] {f.name}: Removed {removed} -> Remaining latest: {wb.sheetnames[-1]}")
            else:
                print(f"[{idx:02d}/{len(files)}] {f.name}: No target sheets found -> Latest: {wb.sheetnames[-1]}")

        except Exception as e:
            print(f"[{idx:02d}/{len(files)}] {f.name}: ERROR: {e}")

    print("\n" + "=" * 60)
    print(f"Completed! Cleaned {len(results)} of {len(files)} files.")
    return results


if __name__ == "__main__":
    import sys

    # Allow custom path via command-line argument if provided
    target_dir = sys.argv[1] if len(sys.argv) > 1 else "/home/im-redmond/Downloads/Station_Reports_Synced_20261118_144442"
    delete_sheets_from_workbooks(target_dir)
