#!/usr/bin/env python3
"""
Sort PDF files from output/ into sort/G1/ and sort/G2/ based on Excel data.
Excel columns: C = ELU (ELO), D = ID, E = Tip opreme (string to clean).
"""

import os
import re
import shutil
import difflib
import pandas as pd

# ========== CONFIGURATION – ADJUST FOR THIS EXCEL ==========
EXCEL_PATH = "data/excel.xlsx"
SHEET_NAME = "G1_ plan usluga"   # or 0 if sheet name changes
SKIP_ROWS = 1                    # skip header row (PPZ | ELU | ID | ...)
COL_ELO = 2                      # column C (0=A,1=B,2=C) – ELU
COL_ID = 3                       # column D – ID
COL_H = 4                        # column E – Tip opreme
DELIMITERS = ['<br>', '\n', '\r\n', ';']
STRINGS_TO_REMOVE = ["Multimedijska oprema - "]
OUTPUT_DIR = "output"
SORT_BASE = "sort"
FUZZY_CUTOFF = 0.6               # for folder matching
# ============================================================

def clean_cell(value):
    if not isinstance(value, str):
        return ""
    value = value.strip()
    if value.startswith("'") or value.startswith('"'):
        value = value[1:]
    if value.endswith("'") or value.endswith('"'):
        value = value[:-1]
    return value.strip()

def load_excel_data(excel_path):
    df = pd.read_excel(excel_path, sheet_name=SHEET_NAME, header=None, dtype=str, engine='openpyxl')
    if SKIP_ROWS > 0:
        df = df.iloc[SKIP_ROWS:]
    rows = []
    for _, row in df.iterrows():
        raw_elo = clean_cell(row[COL_ELO]) if pd.notna(row[COL_ELO]) else ""
        raw_id = clean_cell(row[COL_ID]) if pd.notna(row[COL_ID]) else ""
        raw_h = clean_cell(row[COL_H]) if pd.notna(row[COL_H]) else ""
        if not raw_elo and not raw_id and not raw_h:
            continue
        # Split ELO column if multiple values (e.g., "6a<br>6b")
        elo_list = []
        for delim in DELIMITERS:
            if delim in raw_elo:
                parts = raw_elo.split(delim)
                elo_list = [clean_cell(p) for p in parts if clean_cell(p)]
                break
        if not elo_list:
            elo_list = [raw_elo] if raw_elo else []
        # Clean H
        cleaned_h = raw_h
        for s in STRINGS_TO_REMOVE:
            cleaned_h = cleaned_h.replace(s, "")
        cleaned_h = cleaned_h.strip()
        rows.append((elo_list, raw_id, cleaned_h))
    return rows

def parse_filename(filename):
    pattern = r"^G([12]) ELO ([0-9]+[a-z]?)-([0-9]+)\.pdf$"
    m = re.match(pattern, filename)
    if m:
        return int(m.group(1)), m.group(2), m.group(3)
    return None, None, None

def find_matching_row(rows, file_elo, file_id):
    for elo_list, id_val, cleaned_h in rows:
        if file_elo in elo_list and file_id == id_val:
            return cleaned_h
    # Fallback: substring for ID
    for elo_list, id_val, cleaned_h in rows:
        if file_elo in elo_list and (file_id in id_val or id_val in file_id):
            return cleaned_h
    return None

def find_target_folder(cleaned_name, group):
    base_dir = os.path.join(SORT_BASE, f"G{group}")
    if not os.path.isdir(base_dir):
        return None
    folders = [f for f in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, f))]
    if not folders:
        return None
    
    def normalize(s):
        s = s.lower()
        s = re.sub(r'[^\w\s]', ' ', s)  # replace punctuation with space
        s = re.sub(r'\s+', ' ', s).strip()
        return s
    
    norm_target = normalize(cleaned_name)
    best_match = None
    best_score = 0
    for folder in folders:
        norm_folder = normalize(folder)
        score = difflib.SequenceMatcher(None, norm_target, norm_folder).ratio()
        # Bonus if one contains the other
        if norm_target in norm_folder or norm_folder in norm_target:
            score = max(score, 0.8)
        if score > best_score and score >= FUZZY_CUTOFF:
            best_score = score
            best_match = folder
    if best_match:
        return os.path.join(base_dir, best_match)
    return None

def main():
    if not os.path.exists(EXCEL_PATH):
        print(f"Error: Excel file not found at {EXCEL_PATH}")
        return
    if not os.path.isdir(OUTPUT_DIR):
        print(f"Error: Output directory '{OUTPUT_DIR}' not found")
        return

    rows = load_excel_data(EXCEL_PATH)
    print(f"Loaded {len(rows)} rows from Excel (after skipping {SKIP_ROWS} rows)")

    os.makedirs(SORT_BASE, exist_ok=True)
    for g in ["G1", "G2"]:
        os.makedirs(os.path.join(SORT_BASE, g), exist_ok=True)

    files = [f for f in os.listdir(OUTPUT_DIR) if f.endswith(".pdf")]
    print(f"Found {len(files)} PDF files\n")

    for filename in files:
        group, file_elo, file_id = parse_filename(filename)
        if group is None:
            print(f"SKIP  {filename} (pattern mismatch)")
            continue

        cleaned_h = find_matching_row(rows, file_elo, file_id)
        if cleaned_h is None:
            print(f"FAIL  {filename} -> no Excel match for ELO={file_elo} ID={file_id}")
            continue

        target = find_target_folder(cleaned_h, group)
        if target is None:
            print(f"FAIL  {filename} -> no folder for '{cleaned_h}' in sort/G{group}/")
            continue

        src = os.path.join(OUTPUT_DIR, filename)
        dst = os.path.join(target, filename)
        if os.path.exists(dst):
            print(f"SKIP  {filename} -> already exists")
        else:
            shutil.copy2(src, dst)
            print(f"OK    {filename} -> {target}")

    print("\nDone.")

if __name__ == "__main__":
    main()