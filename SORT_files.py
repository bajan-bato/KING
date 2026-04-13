#!/usr/bin/env python3
"""
Sort PDF files from output/ into sort/G1/ and sort/G2/ based on Excel data.
If Excel lookup fails, fallback to models.csv (which has exact EL/ID from OCR).
Uses fuzzy matching to find the best matching subfolder under sort/G1/ or sort/G2/.
Outputs summary statistics and writes details to log.txt.
Add --debug to write detailed diagnostic info to debug.txt.
"""

import os
import re
import csv
import shutil
import difflib
import argparse
import pandas as pd

# ========== CONFIGURATION ==========
EXCEL_PATH = "data/excel.xlsx"
CSV_PATH = "output/models.csv"
OUTPUT_DIR = "output"
SORT_BASE = "sort"
REMOVE_PREFIX = "Multimedijska oprema - "
LOG_FILE = "log.txt"
DEBUG_FILE = "debug.txt"
# ===================================

def load_tips_from_excel(excel_path):
    tip_lookup = {}
    for group, sheet_name in [(1, "G1_ planiranje"), (2, "G2_ planiranje")]:
        try:
            df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, dtype=str, engine='openpyxl')
            for _, row in df.iterrows():
                if len(row) < 5:
                    continue
                el_cell = str(row[2]) if pd.notna(row[2]) else ""
                id_cell = str(row[3]) if pd.notna(row[3]) else ""
                tip_cell = str(row[4]) if pd.notna(row[4]) else ""
                if not el_cell.strip() or not id_cell.strip() or not tip_cell.strip():
                    continue
                el_list = [el.strip() for el in re.split(r'[\n\r]+', el_cell) if el.strip()]
                for el in el_list:
                    tip_lookup[(group, el, id_cell)] = tip_cell.strip()
        except Exception as e:
            print(f"Warning: Could not read sheet '{sheet_name}': {e}")
    return tip_lookup

def load_tips_from_csv(csv_path):
    tip_lookup = {}
    try:
        with open(csv_path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                el = row.get('EL', '').strip()
                id_val = row.get('ID', '').strip()
                tip = row.get('Tip', row.get('Model', '')).strip()
                if el and id_val and tip:
                    tip_lookup[(el, id_val)] = tip
    except Exception as e:
        print(f"Warning: Could not read {csv_path}: {e}")
    return tip_lookup

def strip_el_suffix(el):
    m = re.match(r'^([0-9]+)[a-z]?$', el)
    return m.group(1) if m else el

def parse_filename(filename):
    pattern = r"^G([12]) ELO ([0-9]+[a-z]?)-([0-9]+)\.pdf$"
    m = re.match(pattern, filename)
    if m:
        return int(m.group(1)), m.group(2), m.group(3)
    return None, None, None

def normalize_name(name):
    name = re.sub(r'^\d+\.\s*G[12]\s+', '', name)
    if name.startswith(REMOVE_PREFIX):
        name = name[len(REMOVE_PREFIX):]
    name = name.lower()
    name = re.sub(r'[^\w\s]', ' ', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name

def find_target_folder(tip, group, sort_base, debug=False, debug_file=None):
    base_dir = os.path.join(sort_base, f"G{group}")
    if not os.path.isdir(base_dir):
        return None, None, None
    folders = [f for f in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, f))]
    if not folders:
        return None, None, None

    norm_tip = normalize_name(tip)
    if debug and debug_file:
        debug_file.write(f"    Normalized tip: '{norm_tip}'\n")
        debug_file.write(f"    Available folders: {folders}\n")

    # 1. Exact match
    for folder in folders:
        norm_folder = normalize_name(folder)
        if norm_tip == norm_folder:
            if debug and debug_file:
                debug_file.write(f"    Exact match found: {folder}\n")
            return os.path.join(base_dir, folder), 'exact', folder

    # 2. Substring match
    for folder in folders:
        norm_folder = normalize_name(folder)
        if norm_tip in norm_folder or norm_folder in norm_tip:
            if debug and debug_file:
                debug_file.write(f"    Substring match found: {folder} (norm_folder='{norm_folder}')\n")
            return os.path.join(base_dir, folder), 'substring', folder

    # 3. Fuzzy ratio
    best_score = -1
    best_matches = []
    for folder in folders:
        norm_folder = normalize_name(folder)
        score = difflib.SequenceMatcher(None, norm_tip, norm_folder).ratio()
        if debug and debug_file:
            debug_file.write(f"    Fuzzy: '{folder}' score = {score:.4f}\n")
        if score > best_score:
            best_score = score
            best_matches = [folder]
        elif score == best_score and score > 0:
            best_matches.append(folder)

    if best_score > 0:
        if len(best_matches) > 1:
            if debug and debug_file:
                debug_file.write(f"    Ambiguous: multiple folders with same score {best_score:.4f}: {best_matches}\n")
            return None, 'ambiguous', f"multiple folders with same score: {', '.join(best_matches)}"
        else:
            if debug and debug_file:
                debug_file.write(f"    Fuzzy match accepted: {best_matches[0]} (score={best_score:.4f})\n")
            return os.path.join(base_dir, best_matches[0]), 'fuzzy', f"score={best_score:.2f}"
    else:
        if debug and debug_file:
            debug_file.write(f"    No match found (best score=0)\n")
        return None, 'no_match', f"no folder found (best score=0)"

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--debug", action="store_true", help="Write detailed diagnostic info to debug.txt")
    args = parser.parse_args()

    debug = args.debug
    debug_file = None
    if debug:
        debug_file = open(DEBUG_FILE, 'w', encoding='utf-8')
        debug_file.write("===== DEBUG SORTING SCRIPT =====\n\n")

    # Load data
    print("Loading Excel tips...")
    excel_lookup = load_tips_from_excel(EXCEL_PATH)
    print(f"Loaded {len(excel_lookup)} entries from Excel.")
    if debug:
        debug_file.write(f"Excel lookup entries: {len(excel_lookup)}\n")
        sample_keys = list(excel_lookup.keys())[:10]
        debug_file.write(f"Sample Excel keys: {sample_keys}\n\n")

    print("Loading CSV fallback...")
    csv_lookup = load_tips_from_csv(CSV_PATH)
    print(f"Loaded {len(csv_lookup)} entries from CSV.")
    if debug:
        debug_file.write(f"CSV lookup entries: {len(csv_lookup)}\n")
        sample_csv = list(csv_lookup.keys())[:10]
        debug_file.write(f"Sample CSV keys: {sample_csv}\n\n")

    if not excel_lookup and not csv_lookup:
        print("No data loaded. Exiting.")
        if debug:
            debug_file.close()
        return

    # Scan output directory
    if not os.path.isdir(OUTPUT_DIR):
        print(f"Error: Output directory '{OUTPUT_DIR}' not found.")
        if debug:
            debug_file.close()
        return
    files = [f for f in os.listdir(OUTPUT_DIR) if f.endswith(".pdf")]
    total_loaded = len(files)
    print(f"Found {total_loaded} PDF files in {OUTPUT_DIR}")

    # Create sort directories
    os.makedirs(SORT_BASE, exist_ok=True)
    for g in ["G1", "G2"]:
        os.makedirs(os.path.join(SORT_BASE, g), exist_ok=True)

    sorted_count = 0
    unsorted_count = 0
    conflict_count = 0
    unsorted_details = []
    conflict_details = []
    not_found_in_csv = []   # store (filename, key) for files that fell back to CSV but not found
    not_found_in_excel = [] # store (filename, key) for files not found in Excel (only if debug)

    for filename in files:
        if debug:
            debug_file.write(f"\n--- Processing: {filename} ---\n")

        group, el, id_val = parse_filename(filename)
        if group is None:
            if debug:
                debug_file.write(f"  Invalid filename format\n")
            unsorted_details.append((filename, "invalid filename format"))
            unsorted_count += 1
            continue

        if debug:
            debug_file.write(f"  Group={group}, EL='{el}', ID='{id_val}'\n")

        tip = None
        source = None

        # 1. Try Excel exact
        key_exact = (group, el, id_val)
        if debug:
            debug_file.write(f"  Trying Excel exact key: {key_exact}\n")
        if key_exact in excel_lookup:
            tip = excel_lookup[key_exact]
            source = "Excel exact"
            if debug:
                debug_file.write(f"    Found! tip='{tip}'\n")
        else:
            if debug:
                debug_file.write(f"    Not found.\n")
            not_found_in_excel.append((filename, key_exact))

        # 2. Try Excel with stripped EL
        if tip is None:
            el_stripped = strip_el_suffix(el)
            if el_stripped != el:
                key_stripped = (group, el_stripped, id_val)
                if debug:
                    debug_file.write(f"  Trying Excel stripped key: {key_stripped}\n")
                if key_stripped in excel_lookup:
                    tip = excel_lookup[key_stripped]
                    source = f"Excel stripped (EL {el} -> {el_stripped})"
                    if debug:
                        debug_file.write(f"    Found! tip='{tip}'\n")
                else:
                    if debug:
                        debug_file.write(f"    Not found.\n")
                    not_found_in_excel.append((filename, key_stripped))

        # 3. Fallback to CSV (ignores group)
        if tip is None:
            key_csv = (el, id_val)
            if debug:
                debug_file.write(f"  Trying CSV key: {key_csv}\n")
            if key_csv in csv_lookup:
                tip = csv_lookup[key_csv]
                source = "CSV fallback"
                if debug:
                    debug_file.write(f"    Found! tip='{tip}'\n")
            else:
                if debug:
                    debug_file.write(f"    Not found.\n")
                not_found_in_csv.append((filename, key_csv))
                # Also try stripped EL in CSV
                el_stripped = strip_el_suffix(el)
                if el_stripped != el:
                    key_csv_stripped = (el_stripped, id_val)
                    if debug:
                        debug_file.write(f"  Trying CSV stripped key: {key_csv_stripped}\n")
                    if key_csv_stripped in csv_lookup:
                        tip = csv_lookup[key_csv_stripped]
                        source = f"CSV stripped (EL {el} -> {el_stripped})"
                        if debug:
                            debug_file.write(f"    Found! tip='{tip}'\n")
                    else:
                        if debug:
                            debug_file.write(f"    Not found.\n")
                        not_found_in_csv.append((filename, key_csv_stripped))

        if tip is None:
            if debug:
                debug_file.write(f"  No tip found.\n")
            unsorted_details.append((filename, f"no tip found for group G{group}, EL={el}, ID={id_val}"))
            unsorted_count += 1
            continue

        if debug:
            debug_file.write(f"  Tip found: '{tip}' (source: {source})\n")

        # Find target folder
        target_dir, match_type, match_info = find_target_folder(tip, group, SORT_BASE, debug=debug, debug_file=debug_file)
        if not target_dir:
            reason = f"no folder match for tip '{tip}' (source: {source})"
            if match_type == 'ambiguous':
                reason += f" - ambiguous: {match_info}"
            elif match_type == 'no_match':
                reason += f" - {match_info}"
            if debug:
                debug_file.write(f"  Folder match failed: {reason}\n")
            unsorted_details.append((filename, reason))
            unsorted_count += 1
            continue

        src = os.path.join(OUTPUT_DIR, filename)
        dst = os.path.join(target_dir, filename)
        if os.path.exists(dst):
            if debug:
                debug_file.write(f"  Conflict: file already exists in {os.path.basename(target_dir)}\n")
            conflict_details.append((filename, os.path.basename(target_dir)))
            conflict_count += 1
            continue

        shutil.copy2(src, dst)
        sorted_count += 1
        print(f"OK    {filename} -> {os.path.basename(target_dir)} (match: {match_type}, source: {source})")
        if debug:
            debug_file.write(f"  Sorted to {os.path.basename(target_dir)}\n")

    # Summary
    print("\n" + "="*50)
    print("SUMMARY")
    print("="*50)
    print(f"[LOADED]:   {total_loaded}")
    print(f"[SORTED]:   {sorted_count}")
    print(f"[UNSORTED]: {unsorted_count}")
    print(f"[CONFLICTS]: {conflict_count}")
    print(f"\nDetails written to {LOG_FILE}")
    if debug:
        print(f"Debug info written to {DEBUG_FILE}")

    # Write log
    with open(LOG_FILE, 'w', encoding='utf-8') as log:
        if unsorted_details:
            log.write("=========\nUNSORTED\n=========\n")
            for fname, reason in unsorted_details:
                log.write(f"{fname}: {reason}\n")
            log.write("\n")
        if conflict_details:
            log.write("=========\nCONFLICTS\n=========\n")
            for fname, target_folder in conflict_details:
                log.write(f"{fname} (already exists in {target_folder})\n")
            log.write("\n")
        if not unsorted_details and not conflict_details:
            log.write("No issues.\n")

    if debug:
        debug_file.write("\n" + "="*50 + "\n")
        debug_file.write("FILES NOT FOUND IN CSV (after Excel fallback):\n")
        if not_found_in_csv:
            for fname, key in not_found_in_csv:
                debug_file.write(f"  {fname}: key={key}\n")
        else:
            debug_file.write("  None\n")
        debug_file.write("\nFILES NOT FOUND IN EXCEL (exact or stripped):\n")
        if not_found_in_excel:
            for fname, key in not_found_in_excel:
                debug_file.write(f"  {fname}: key={key}\n")
        else:
            debug_file.write("  None\n")
        debug_file.close()

    print("\nDone.")

if __name__ == "__main__":
    main()