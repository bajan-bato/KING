#!/usr/bin/env python3
"""
Sort PDF files from output/ into sort/G1/ and sort/G2/ based on Excel data.
Always returns the best matching folder (no cutoff).
Minimal logging: OK / FAIL / SKIP only.
"""

import os
import re
import shutil
import difflib
import pandas as pd

# ========== CONFIGURATION ==========
EXCEL_PATH = "data/excel.xlsx"
SHEET_NAME = "30.3.G2"
SKIP_ROWS = 1
COL_ELO = 1
COL_ID = 2
COL_H = 3
DELIMITERS = ['<br>', '\n', '\r\n', ';']
STRINGS_TO_REMOVE = ["Multimedijska oprema - "]
OUTPUT_DIR = "output"
SORT_BASE = "sort"
# ===================================

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
        elo_list = []
        for delim in DELIMITERS:
            if delim in raw_elo:
                parts = raw_elo.split(delim)
                elo_list = [clean_cell(p) for p in parts if clean_cell(p)]
                break
        if not elo_list:
            elo_list = [raw_elo] if raw_elo else []
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
    # Exact ID match first
    for elo_list, id_val, cleaned_h in rows:
        if file_elo in elo_list and file_id == id_val:
            return cleaned_h
    # Substring ID match
    for elo_list, id_val, cleaned_h in rows:
        if file_elo in elo_list and (file_id in id_val or id_val in file_id):
            return cleaned_h
    # ELO-only fallback (ignore ID)
    for elo_list, id_val, cleaned_h in rows:
        if file_elo in elo_list:
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
        s = re.sub(r'^\d+\.\s*g[12]\s+', '', s)   # remove "1. G1 " prefix
        s = re.sub(r'[^\w\s]', ' ', s)
        s = re.sub(r'\s+', ' ', s).strip()
        return s

    norm_cleaned = normalize(cleaned_name)

    # FIRST: try substring match (contains) – most reliable
    best_substring_match = None
    for folder in folders:
        norm_folder = normalize(folder)
        if norm_cleaned in norm_folder or norm_folder in norm_cleaned:
            # If multiple, choose the one with the longer match (or first)
            if best_substring_match is None:
                best_substring_match = folder
            else:
                # Keep the one where the match is longer (optional)
                pass
    if best_substring_match:
        return os.path.join(base_dir, best_substring_match)

    # SECOND: fallback to fuzzy ratio (only if no substring match)
    best_match = None
    best_score = -1
    for folder in folders:
        norm_folder = normalize(folder)
        score = difflib.SequenceMatcher(None, norm_cleaned, norm_folder).ratio()
        if score > best_score:
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
    print(f"Loaded {len(rows)} rows from Excel")

    os.makedirs(SORT_BASE, exist_ok=True)
    for g in ["G1", "G2"]:
        os.makedirs(os.path.join(SORT_BASE, g), exist_ok=True)

    files = [f for f in os.listdir(OUTPUT_DIR) if f.endswith(".pdf")]
    print(f"Found {len(files)} PDF files\n")

    for filename in files:
        group, file_elo, file_id = parse_filename(filename)
        if group is None:
            print(f"SKIP  {filename}")
            continue

        cleaned_h = find_matching_row(rows, file_elo, file_id)
        if cleaned_h is None:
            print(f"FAIL  {filename} (no Excel match)")
            continue

        target = find_target_folder(cleaned_h, group)
        if target is None:
            print(f"FAIL  {filename} (no folder match)")
            continue

        src = os.path.join(OUTPUT_DIR, filename)
        dst = os.path.join(target, filename)
        if os.path.exists(dst):
            print(f"SKIP  {filename} (already exists)")
        else:
            shutil.copy2(src, dst)
            print(f"OK    {filename} -> {os.path.basename(target)}")

    print("\nDone.")

if __name__ == "__main__":
    main()