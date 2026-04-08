#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import glob
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from pathlib import Path

# ==================== CONFIGURATION ====================
DATA_DIR = "data"
IN_DIR = "in"
OUT_DIR = "out"
EXCEL_FILE = os.path.join(DATA_DIR, "excel.xlsx")
SERIJSKI_FILE = os.path.join(DATA_DIR, "serijski.xlsx")

# Create output directory if it doesn't exist
Path(OUT_DIR).mkdir(parents=True, exist_ok=True)

# ==================== HELPER FUNCTIONS ====================
def normalize_string(s):
    """Lowercase and strip whitespace, convert to string."""
    return str(s).strip().lower() if pd.notna(s) else ""

def set_cell_text(cell, text, font_name='Arial', font_size=11):
    """Set cell text with specified font (Arial, size 11)."""
    cell.text = ''
    paragraph = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    run = paragraph.add_run(str(text) if text is not None else '')
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

def split_and_pair(el_str, code_str):
    """
    Split el_str by newline, split code_str by newline, clean each code line
    by taking the first token before any space, then pair according to rules:
    - If only one code, pair it with every el.
    - If multiple codes, pair one-to-one (truncate to shorter list).
    Returns list of (el, code) tuples.
    """
    # Split and clean EL
    el_list = [e.strip() for e in str(el_str).split('\n') if e.strip()]
    
    # Split and clean CODE: each line -> first token before whitespace
    raw_codes = [c.strip() for c in str(code_str).split('\n') if c.strip()]
    code_list = []
    for raw in raw_codes:
        parts = raw.split()  # split by any whitespace
        if parts:
            code_list.append(parts[0])  # take the first token (the actual code)
    
    if not code_list:
        return []
    if len(code_list) == 1:
        code_single = code_list[0]
        return [(el, code_single) for el in el_list]
    # Multiple codes: one-to-one
    pairs = []
    for i in range(min(len(el_list), len(code_list))):
        pairs.append((el_list[i], code_list[i]))
    return pairs

# ==================== 1. READ EXCEL FILES ====================
try:
    df_excel = pd.read_excel(EXCEL_FILE, sheet_name="Sheet1", header=None, dtype=str, engine='openpyxl')
except Exception as e:
    print(f"Error reading {EXCEL_FILE}: {e}")
    print("Make sure openpyxl is installed: pip install openpyxl")
    exit(1)

# Explode rows: each original row may produce multiple (el, id, code) rows
exploded_rows = []
for idx, row in df_excel.iterrows():
    # Columns: B -> index 1, C -> index 2, D -> index 3
    el_raw = row[1] if pd.notna(row[1]) else ""
    id_raw = row[2] if pd.notna(row[2]) else ""
    code_raw = row[3] if pd.notna(row[3]) else ""

    if not id_raw:  # column C (ID) must be present
        continue

    # Get pairs of (el, code)
    pairs = split_and_pair(el_raw, code_raw)
    for el, code in pairs:
        if el and code:
            exploded_rows.append((normalize_string(el), normalize_string(id_raw), code.strip()))

if not exploded_rows:
    print("No valid rows after explosion. Check excel.xlsx content.")
    exit(1)

# Read serijski.xlsx – columns C (index 2) and F (index 5)
try:
    df_ser = pd.read_excel(SERIJSKI_FILE, sheet_name="Sheet1", header=None, dtype=str, engine='openpyxl')
except Exception as e:
    print(f"Error reading {SERIJSKI_FILE}: {e}")
    exit(1)

serijski_map = {}  # key = normalized C (with padded zeros), value = list of normalized F
for idx, row in df_ser.iterrows():
    c_val = row[2] if pd.notna(row[2]) else ""
    f_val = row[5] if pd.notna(row[5]) else ""
    if c_val and f_val:
        key = normalize_string(c_val)
        serijski_map.setdefault(key, []).append(f_val.strip())

# ==================== 2. PROCESS EACH EXPLODED ROW ====================
summary_data = []  # for output xlsx (el, id, code, serijski_broj)

for el, id_, code in exploded_rows:
    # Build padded code: add "00" at the beginning
    padded_code = "00" + code
    normalized_padded = normalize_string(padded_code)

    # Find matching serijski entries
    matching_serijski = serijski_map.get(normalized_padded, [])

    if not matching_serijski:
        print(f"⚠️ No matching serijski for code {code} (padded: {padded_code})")
        continue

    # ==================== 3. FIND MATCHING .DOCX FILE ====================
    docx_files = glob.glob(os.path.join(IN_DIR, "*.docx"))
    matching_docx = None
    for fpath in docx_files:
        fname = os.path.basename(fpath).lower()
        if el in fname and id_ in fname:
            matching_docx = fpath
            break
    if not matching_docx:
        print(f"⚠️ No .docx file found for el={el}, id={id_}")
        continue

    # ==================== 4. MODIFY THE DOCX ====================
    doc = Document(matching_docx)
    tables = doc.tables
    if len(tables) < 2:
        print(f"⚠️ File {matching_docx} has less than 2 tables. Skipping.")
        continue
    second_table = tables[1]

    # Find column index of "Serijski broj"
    header_row = second_table.rows[0]
    serijski_col_idx = None
    for idx, cell in enumerate(header_row.cells):
        if normalize_string(cell.text) == "serijski broj":
            serijski_col_idx = idx
            break
    if serijski_col_idx is None:
        print(f"⚠️ Column 'Serijski broj' not found in second table of {matching_docx}")
        continue

    # Write serijski brojevi into rows (skip header row)
    num_data_rows = len(second_table.rows) - 1
    for row_idx in range(num_data_rows):
        if row_idx < len(matching_serijski):
            serijski_value = matching_serijski[row_idx]
            cell = second_table.rows[row_idx + 1].cells[serijski_col_idx]
            set_cell_text(cell, serijski_value)

    # Save modified document
    out_filename = f"G1 ELO {el}-{id_}.docx"
    out_path = os.path.join(OUT_DIR, out_filename)
    doc.save(out_path)
    print(f"✅ Saved: {out_path}")

    # Collect summary data
    for serijski_broj in matching_serijski:
        summary_data.append({
            "A": el,
            "B": id_,
            "C": code,
            "D": serijski_broj
        })

# ==================== 5. CREATE SUMMARY XLSX ====================
if summary_data:
    df_summary = pd.DataFrame(summary_data)
    summary_path = os.path.join(OUT_DIR, "joined_data.xlsx")
    df_summary.to_excel(summary_path, index=False, engine='openpyxl')
    print(f"📊 Summary saved: {summary_path}")
else:
    print("No data to write to summary xlsx.")