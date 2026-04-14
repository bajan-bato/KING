#!/usr/bin/env python3
"""
Match EL/ID/otpremnica from first Excel, join with otpremnica/serijski from second Excel,
fill Word templates (second table, column "Serijski broj") with serial numbers.
Outputs joined.xlsx and flagged rows.
"""

import os
import re
import shutil
import pandas as pd
from pathlib import Path
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

# ========== CONFIGURATION ==========
# First Excel (EL, ID, otpremnica)
EXCEL1_PATH = "data/excel.xlsx"
EXCEL1_SHEET = "Sheet1"
EXCEL1_COL_EL = 1
EXCEL1_COL_ID = 2
EXCEL1_COL_OTP = 3

# Second Excel (otpremnica with "00" prefix, serijski)
EXCEL2_PATH = "data/serijski.xlsx"
EXCEL2_SHEET = "Sheet1"
EXCEL2_COL_OTP = 2
EXCEL2_COL_SER = 5

INPUT_DIR = "in"
OUTPUT_DIR = "out"
# ===================================

def load_first_excel():
    # Use openpyxl engine for .xlsx files
    try:
        df = pd.read_excel(EXCEL1_PATH, sheet_name=EXCEL1_SHEET, header=None, dtype=str, engine='openpyxl')
    except Exception as e:
        print(f"Error reading first Excel: {e}")
        print("Make sure openpyxl is installed: pip install openpyxl")
        return [], []
    mapping = []
    flagged_rows = []
    for idx, row in df.iterrows():
        el_cell = row[EXCEL1_COL_EL] if len(row) > EXCEL1_COL_EL else None
        id_cell = row[EXCEL1_COL_ID] if len(row) > EXCEL1_COL_ID else None
        otp_cell = row[EXCEL1_COL_OTP] if len(row) > EXCEL1_COL_OTP else None
        if pd.isna(id_cell):
            continue
        id_val = str(id_cell).strip()
        pairs, flagged = parse_el_otp_pair(el_cell, otp_cell)
        if flagged:
            flagged_rows.append((idx+1, el_cell, id_val, otp_cell))
        for el, otp in pairs:
            mapping.append((el, id_val, otp))
    return mapping, flagged_rows

def load_second_excel():
    try:
        df = pd.read_excel(EXCEL2_PATH, sheet_name=EXCEL2_SHEET, header=None, dtype=str, engine='openpyxl')
    except Exception as e:
        print(f"Error reading second Excel: {e}")
        print("Make sure openpyxl is installed: pip install openpyxl")
        return {}
    otp_to_ser = {}
    for _, row in df.iterrows():
        otp_cell = row[EXCEL2_COL_OTP] if len(row) > EXCEL2_COL_OTP else None
        ser_cell = row[EXCEL2_COL_SER] if len(row) > EXCEL2_COL_SER else None
        if pd.isna(otp_cell) or pd.isna(ser_cell):
            continue
        otp_raw = str(otp_cell).strip()
        # Remove leading zeros (two zeros expected, but remove all)
        otp_clean = otp_raw.lstrip('0')
        ser = str(ser_cell).strip()
        if otp_clean not in otp_to_ser:
            otp_to_ser[otp_clean] = []
        otp_to_ser[otp_clean].append(ser)
    return otp_to_ser

def parse_el_otp_pair(el_cell, otp_cell):
    if pd.isna(el_cell) or pd.isna(otp_cell):
        return [], False
    el_str = str(el_cell).strip()
    otp_str = str(otp_cell).strip()
    el_list = [e.strip() for e in re.split(r'[\n\r]+', el_str) if e.strip()]
    # Split otp by newline, then extract only the first token (the number)
    otp_lines = [line.strip() for line in re.split(r'[\n\r]+', otp_str) if line.strip()]
    otp_numbers = []
    for line in otp_lines:
        # Extract first word (before space) – assume it's the otpremnica number
        parts = line.split()
        if parts:
            otp_numbers.append(parts[0].strip())
        else:
            otp_numbers.append('')
    otp_list = [o for o in otp_numbers if o]
    if not el_list or not otp_list:
        return [], False
    if len(el_list) == 1 and len(otp_list) == 1:
        return [(el_list[0], otp_list[0])], False
    if len(el_list) == len(otp_list):
        return [(el_list[i], otp_list[i]) for i in range(len(el_list))], False
    return [], True

def parse_docx_filename(filename):
    """Extract group, el, id from filename like 'G1 ELO 123-456.docx'"""
    pattern = r'^G([12])\s+ELO\s+([0-9]+[a-z]?)\s*[-–]\s*([0-9]+)\.docx$'
    m = re.match(pattern, filename, re.IGNORECASE)
    if m:
        return int(m.group(1)), m.group(2), m.group(3)
    return None, None, None

def set_cell_text(cell, text, font_name='Arial', font_size=11):
    cell.text = ''
    paragraph = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    run = paragraph.add_run(str(text) if text is not None else '')
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

def clone_row(table, source_row_idx):
    """Clone a table row (preserve formatting)."""
    from copy import deepcopy
    source_row = table.rows[source_row_idx]
    tr = source_row._tr
    new_tr = deepcopy(tr)
    table._tbl.append(new_tr)
    return table.rows[-1]

def fill_serial_numbers(doc, serijski_list):
    """Find second table, locate 'Serijski broj' column, fill rows."""
    if len(doc.tables) < 2:
        print("  Warning: Document has fewer than 2 tables.")
        return False
    table = doc.tables[1]
    # Find header row (assume first row)
    header = table.rows[0]
    ser_col_idx = None
    for idx, cell in enumerate(header.cells):
        if "serijski broj" in cell.text.strip().lower():
            ser_col_idx = idx
            break
    if ser_col_idx is None:
        print("  Warning: Column 'Serijski broj' not found.")
        return False
    # Find first data row (row index 1)
    data_start_idx = 1
    if len(table.rows) <= data_start_idx:
        table.add_row()
    # Count needed rows
    needed = len(serijski_list)
    current_data_rows = len(table.rows) - 1
    if needed > current_data_rows:
        for _ in range(needed - current_data_rows):
            clone_row(table, data_start_idx)
    # Write serial numbers
    for i, ser in enumerate(serijski_list):
        row = table.rows[data_start_idx + i]
        set_cell_text(row.cells[ser_col_idx], ser)
    return True

def main():
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)

    # Step 1: Load first Excel
    print("Loading first Excel...")
    el_id_otp, flagged_rows = load_first_excel()
    print(f"  Loaded {len(el_id_otp)} (EL, ID, otpremnica) entries.")
    if flagged_rows:
        print(f"  Flagged {len(flagged_rows)} rows (mismatched EL/otpremnica counts).")

    # Step 2: Load second Excel
    print("\nLoading second Excel...")
    otp_to_ser = load_second_excel()
    print(f"  Loaded {len(otp_to_ser)} unique otpremnice with serials.")

    # Step 3: Join
    print("\nJoining data...")
    joined = []
    missing_otp = []
    for el, id_val, otp in el_id_otp:
        ser_list = otp_to_ser.get(otp, [])
        if not ser_list:
            missing_otp.append((el, id_val, otp))
        for ser in ser_list:
            joined.append((el, id_val, otp, ser))
    print(f"  Joined {len(joined)} rows (EL, ID, otpremnica, serijski).")
    if missing_otp:
        print(f"  Warning: {len(missing_otp)} otpremnice not found in second Excel.")

    # Step 4: Save joined.xlsx
    joined_df = pd.DataFrame(joined, columns=['EL', 'ID', 'Otpremnica', 'Serijski'])
    joined_path = os.path.join(OUTPUT_DIR, "joined.xlsx")
    joined_df.to_excel(joined_path, index=False, engine='openpyxl')
    print(f"\nSaved joined data to {joined_path}")

    # Step 5: Process DOCX files
    print("\nProcessing DOCX files...")
    if not os.path.isdir(INPUT_DIR):
        print(f"  Input directory '{INPUT_DIR}' not found, skipping.")
    else:
        docx_files = [f for f in os.listdir(INPUT_DIR) if f.endswith('.docx')]
        print(f"  Found {len(docx_files)} DOCX files.")
        from collections import defaultdict
        el_id_to_ser = defaultdict(list)
        for el, id_val, otp, ser in joined:
            el_id_to_ser[(el, id_val)].append(ser)
        processed = 0
        for filename in docx_files:
            group, el, id_val = parse_docx_filename(filename)
            if group is None:
                print(f"  SKIP: {filename} (invalid name format)")
                continue
            ser_list = el_id_to_ser.get((el, id_val), [])
            if not ser_list:
                print(f"  SKIP: {filename} (no serial numbers found for EL={el}, ID={id_val})")
                continue
            doc_path = os.path.join(INPUT_DIR, filename)
            doc = Document(doc_path)
            if fill_serial_numbers(doc, ser_list):
                out_name = f"G{group} ELO {el}-{id_val}.docx"
                out_path = os.path.join(OUTPUT_DIR, out_name)
                doc.save(out_path)
                print(f"  OK: {out_name} ({len(ser_list)} serial(s))")
                processed += 1
            else:
                print(f"  FAIL: {filename} (table error)")
        print(f"  Processed {processed} DOCX files.")

    # Step 6: Output flagged rows from first Excel
    if flagged_rows:
        print("\n" + "="*60)
        print("FLAGGED ROWS FROM FIRST EXCEL (mismatched EL/otpremnica counts):")
        for row_num, el, id_val, otp in flagged_rows:
            print(f"  Row {row_num}: EL='{el}', ID='{id_val}', OTP='{otp}'")
        print("="*60)

    # Optionally save flagged to CSV
    if flagged_rows:
        flagged_df = pd.DataFrame(flagged_rows, columns=['Row', 'EL', 'ID', 'Otpremnica'])
        flagged_path = os.path.join(OUTPUT_DIR, "flagged_rows.xlsx")
        flagged_df.to_excel(flagged_path, index=False, engine='openpyxl')
        print(f"\nFlagged rows saved to {flagged_path}")

    print("\nDone.")

if __name__ == "__main__":
    main()