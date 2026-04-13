#!/usr/bin/env python3
"""
Count individual ELs marked as '1' or '2' in the G1_ planiranje and G2_ planiranje sheets.
For each row:
- If column B is exactly "1" or "2": count all ELs in column C (split by newline).
- If column B contains "Samo ... 1" or "Samo ... 2": count only the ELs listed after "Samo".
Outputs total counts per sheet and overall.
"""

import re
import pandas as pd
from pathlib import Path

# ========== CONFIGURATION ==========
EXCEL_PATH = "out/updated_excel.xlsx"   # or wherever your updated file is
OUTPUT_FILE = "count_summary.txt"       # optional
# ===================================

def count_els_in_row(cell_c, cell_b):
    """Return (count_for_1, count_for_2) based on cell B and C."""
    if pd.isna(cell_c) or not str(cell_c).strip():
        return 0, 0
    c_val = str(cell_c).strip()
    b_val = str(cell_b).strip() if not pd.isna(cell_b) else ""
    
    # Split C into individual ELs (newline or other delimiters? Assume newline)
    # But also handle cases where C might contain commas? Use newline primarily.
    el_list = [el.strip() for el in re.split(r'[\n\r]+', c_val) if el.strip()]
    if not el_list:
        return 0, 0
    
    # Determine what to count based on B
    if b_val == "1":
        return len(el_list), 0
    elif b_val == "2":
        return 0, len(el_list)
    elif b_val.startswith("Samo"):
        # Extract the list of ELs between "Samo " and the trailing number
        # Pattern: "Samo 300a, 300b 1" or "Samo 300a 1" (no comma)
        # We'll split by comma and space, but simpler: find the part before the last space
        parts = b_val.split()
        # parts like ['Samo', '300a,', '300b', '1'] or ['Samo', '300a,300b', '1']
        # Better: use regex to capture the list between "Samo " and the final number
        match = re.search(r'Samo\s+(.+?)\s+([12])$', b_val)
        if match:
            list_str = match.group(1)
            target_num = match.group(2)
            # Split list by commas and spaces
            listed_els = [el.strip().rstrip(',') for el in re.split(r'[,\s]+', list_str) if el.strip()]
            # Count how many of these listed ELs actually exist in C? The user wants to count each listed EL separately.
            # But to be safe, we count only those that are in el_list? The description says "count every XXXz separately" from the list.
            # We'll count all listed ELs (assuming they are valid). No need to filter against C because the list comes from the match.
            count = len(listed_els)
            if target_num == "1":
                return count, 0
            else:
                return 0, count
        else:
            # Fallback: old format maybe "Samo 300a, 300b 1" without space after comma? Try split on comma
            if ',' in b_val:
                # extract everything between 'Samo ' and last space
                last_space = b_val.rfind(' ')
                if last_space > 5:
                    list_part = b_val[5:last_space]  # after 'Samo '
                    listed = [x.strip() for x in list_part.split(',')]
                    target_num = b_val[last_space+1:]
                    if target_num == "1":
                        return len(listed), 0
                    else:
                        return 0, len(listed)
    # If B is something else (e.g., empty or unknown), return 0
    return 0, 0

def process_sheet(sheet_name, df):
    """Process one sheet and print counts."""
    total_1 = 0
    total_2 = 0
    rows_processed = 0
    for idx, row in df.iterrows():
        # Column C is index 2, column B is index 1
        c_val = row.iloc[2] if len(row) > 2 else None
        b_val = row.iloc[1] if len(row) > 1 else None
        cnt1, cnt2 = count_els_in_row(c_val, b_val)
        if cnt1 or cnt2:
            rows_processed += 1
            total_1 += cnt1
            total_2 += cnt2
    print(f"\nSheet: {sheet_name}")
    print(f"  Rows with matches: {rows_processed}")
    print(f"  Total ELs counted as '1': {total_1}")
    print(f"  Total ELs counted as '2': {total_2}")
    return total_1, total_2

def main():
    if not Path(EXCEL_PATH).exists():
        print(f"Error: Excel file not found at {EXCEL_PATH}")
        return
    
    # Read both sheets
    try:
        df_g1 = pd.read_excel(EXCEL_PATH, sheet_name="G1_ planiranje", header=None, dtype=str)
        df_g2 = pd.read_excel(EXCEL_PATH, sheet_name="G2_ planiranje", header=None, dtype=str)
    except Exception as e:
        print(f"Error reading Excel: {e}")
        return
    
    print("="*60)
    print("COUNTING ELs MARKED AS '1' OR '2'")
    print("="*60)
    
    g1_1, g1_2 = process_sheet("G1_ planiranje", df_g1)
    g2_1, g2_2 = process_sheet("G2_ planiranje", df_g2)
    
    print("\n" + "="*60)
    print("TOTALS")
    print("="*60)
    print(f"Overall ELs counted as '1': {g1_1 + g2_1}")
    print(f"Overall ELs counted as '2': {g1_2 + g2_2}")
    print(f"Total ELs: {g1_1 + g2_1 + g1_2 + g2_2}")
    
    # Optional: save summary to file
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write("COUNT SUMMARY\n")
        f.write("="*40 + "\n")
        f.write(f"G1_ planiranje: 1's = {g1_1}, 2's = {g1_2}\n")
        f.write(f"G2_ planiranje: 1's = {g2_1}, 2's = {g2_2}\n")
        f.write(f"Overall: 1's = {g1_1+g2_1}, 2's = {g1_2+g2_2}\n")
    print(f"\nSummary saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    main()