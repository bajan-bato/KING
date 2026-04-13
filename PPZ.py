#!/usr/bin/env python3
"""
Scan all PDF files in sort/ (including subfolders), compare with Excel sheets
"G1_ planiranje" and "G2_ planiranje" from data/excel.xlsx.
For each row, if all EL values (split by newline) for a given ID are present in the files,
write "1" in column B. If only some are present, write "Samo <list> 1".
If none are present, leave column B unchanged (do NOT write anything).
Output is a copy of the original Excel file (preserving all formatting) with column B updated only for matches.
Logs only matched rows (1 or partial) and prints unmatched files at the end.
"""

import os
import re
import shutil
from pathlib import Path
from openpyxl import load_workbook

# ========== CONFIGURATION ==========
SORT_BASE = "sort"
EXCEL_PATH = "data/CARNET_G1_G2_plan_v0.15.xlsx"
OUTPUT_DIR = "out"
OUTPUT_FILE = "updated_excel.xlsx"
# ===================================

def get_files_by_group():
    """Return dict {(group, el, id): full_path} for all PDFs in sort/."""
    pattern = re.compile(r'^G([12]) ELO ([0-9]+[a-z]?)-([0-9]+)\.pdf$')
    file_dict = {}
    for root, dirs, files in os.walk(SORT_BASE):
        for f in files:
            m = pattern.match(f)
            if m:
                group = int(m.group(1))
                el = m.group(2)
                id_val = m.group(3)
                full_path = os.path.join(root, f)
                file_dict[(group, el, id_val)] = full_path
    return file_dict

def process_sheet(wb, sheet_name, file_dict, group_num):
    """Update column B only for rows that have a match. Return set of matched keys."""
    if sheet_name not in wb.sheetnames:
        print(f"  Sheet '{sheet_name}' not found, skipping.")
        return set()
    ws = wb[sheet_name]
    matched_keys = set()

    max_row = ws.max_row
    for row_idx in range(1, max_row + 1):
        # Column C (3) = EL, column D (4) = ID
        el_cell = ws.cell(row=row_idx, column=3)
        id_cell = ws.cell(row=row_idx, column=4)
        el_val = str(el_cell.value) if el_cell.value is not None else ""
        id_val = str(id_cell.value) if id_cell.value is not None else ""
        if not el_val.strip() or not id_val.strip():
            continue

        # Split EL cell by newline
        el_list = [el.strip() for el in re.split(r'[\n\r]+', el_val) if el.strip()]
        if not el_list:
            continue

        present = []
        for el in el_list:
            key = (group_num, el, id_val)
            if key in file_dict:
                present.append(el)
                matched_keys.add(key)

        if not present:
            # No match – leave column B unchanged
            continue

        if len(present) == len(el_list):
            output = "1"
            print(f"  Row {row_idx}: all ELs present -> {output} (ID {id_val})")
        else:
            output = f"Samo {', '.join(present)} 1"
            print(f"  Row {row_idx}: only {', '.join(present)} -> {output} (ID {id_val})")

        # Write to column B (index 2)
        ws.cell(row=row_idx, column=2, value=output)

    return matched_keys

def main():
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)

    # Copy original Excel to output (preserve all formatting)
    output_path = os.path.join(OUTPUT_DIR, OUTPUT_FILE)
    shutil.copy2(EXCEL_PATH, output_path)
    print(f"Copied original Excel file to {output_path}")

    # Scan sort/ folder
    print("Scanning sort/ directory for PDF files...")
    file_dict = get_files_by_group()
    print(f"Found {len(file_dict)} matching PDF files.\n")

    # Load workbook and process sheets
    wb = load_workbook(output_path)
    all_matched_keys = set()

    for group, sheet_name in [(1, "G1_ planiranje"), (2, "G2_ planiranje")]:
        print(f"Processing sheet '{sheet_name}'...")
        matched = process_sheet(wb, sheet_name, file_dict, group)
        all_matched_keys.update(matched)
        print(f"  Done.\n")

    wb.save(output_path)
    print(f"Saved updated workbook to {output_path}")

    # Report unmatched files
    all_keys = set(file_dict.keys())
    unmatched = all_keys - all_matched_keys
    if unmatched:
        print("\n" + "="*60)
        print("UNMATCHED FILES (in sort/ but not matched in any Excel row):")
        for key in sorted(unmatched):
            group, el, id_val = key
            path = file_dict[key]
            print(f"  G{group} ELO {el}-{id_val}.pdf -> {path}")
        print("="*60)
    else:
        print("\nAll PDF files were matched to an Excel row.")

    print("\nDone.")

if __name__ == "__main__":
    main()