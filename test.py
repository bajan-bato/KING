#!/usr/bin/env python3
"""
Compare column B (index 2) between two Excel files for sheets G1_ planiranje and G2_ planiranje.
Matches rows by EL (col C) and ID (col D).
Outputs a report of differences (old vs new) to console and optionally to a CSV file.
Usage:
    python compare_excel_b.py file1.xlsx file2.xlsx
"""

import sys
import argparse
import pandas as pd
from pathlib import Path

def load_sheet_data(file_path, sheet_name):
    """
    Load sheet and return a dict {(el, id): value_in_col_B}.
    Skips rows where EL or ID is missing.
    """
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, dtype=str, engine='openpyxl')
    data = {}
    for idx, row in df.iterrows():
        if len(row) < 4:
            continue
        el = str(row[2]).strip() if pd.notna(row[2]) else ""
        id_val = str(row[3]).strip() if pd.notna(row[3]) else ""
        if not el or not id_val:
            continue
        b_val = str(row[1]).strip() if pd.notna(row[1]) else ""
        data[(el, id_val)] = b_val
    return data

def compare_sheet(sheet_name, old_data, new_data):
    """Compare two dicts and print differences."""
    all_keys = set(old_data.keys()) | set(new_data.keys())
    differences = []
    for key in sorted(all_keys):
        old_val = old_data.get(key, "<MISSING>")
        new_val = new_data.get(key, "<MISSING>")
        if old_val != new_val:
            differences.append((key[0], key[1], old_val, new_val))
    return differences

def main():
    parser = argparse.ArgumentParser(description="Compare column B between two Excel files (G1 and G2 sheets).")
    parser.add_argument("file1", help="First Excel file (e.g., old version)")
    parser.add_argument("file2", help="Second Excel file (e.g., new version)")
    parser.add_argument("--output", help="Save differences to a CSV file (optional)")
    args = parser.parse_args()

    if not Path(args.file1).exists():
        print(f"Error: {args.file1} not found.")
        sys.exit(1)
    if not Path(args.file2).exists():
        print(f"Error: {args.file2} not found.")
        sys.exit(1)

    sheets = ["G1_ planiranje", "G2_ planiranje"]
    all_diffs = []

    for sheet in sheets:
        print(f"\n--- Sheet: {sheet} ---")
        try:
            old_data = load_sheet_data(args.file1, sheet)
            new_data = load_sheet_data(args.file2, sheet)
        except Exception as e:
            print(f"Error reading sheet {sheet}: {e}")
            continue

        diffs = compare_sheet(sheet, old_data, new_data)
        if diffs:
            print(f"Found {len(diffs)} differences:")
            for el, id_val, old, new in diffs:
                print(f"  EL={el}, ID={id_val}: '{old}' -> '{new}'")
                all_diffs.append((sheet, el, id_val, old, new))
        else:
            print("  No differences found.")

    if all_diffs and args.output:
        import csv
        with open(args.output, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(["Sheet", "EL", "ID", "Old Value", "New Value"])
            writer.writerows(all_diffs)
        print(f"\nDifferences saved to {args.output}")

if __name__ == "__main__":
    main()