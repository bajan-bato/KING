#!/usr/bin/env python3
"""
Split scanned PDF(s) into groups using OCR.space API.
Finds "Evidencijska lista XXX-YYYY" and "Grupa 1/2", groups pages accordingly.
Looks up model from Excel (columns C/D for EL/ID, column E for model).
Adds a user‑defined number to each CSV row.
Can append to existing CSV.
Flags files with detailed reasons:
  - single page
  - >2 pages but only 1 marker
  - duplicate EL-ID inside the same file (shows which EL-ID)
  - duplicate EL-ID across different files (shows EL-ID and previous file)
"""

import os
import re
import sys
import csv
import time
import argparse
import tempfile
from io import BytesIO
from typing import List, Tuple, Optional, Dict, Set

import requests
import pandas as pd
from PIL import Image
from pdf2image import convert_from_path
from PyPDF2 import PdfReader, PdfWriter

# ========== API CONFIGURATION ==========
API_KEY = "K84455376988957"
OCR_API_URL = "https://api.ocr.space/parse/image"
OCR_ENGINE = 2
LANGUAGE = ""
MAX_RETRIES = 3
RETRY_DELAY = 5

request_count = 0

# ========== EXCEL CONFIGURATION ==========
EXCEL_PATH = "data/excel.xlsx"   # change this variable as needed
# =========================================

def get_image_size_kb(img: Image.Image, fmt='JPEG', quality=85) -> float:
    buf = BytesIO()
    img.save(buf, format=fmt, quality=quality, optimize=True)
    size = buf.tell() / 1024
    buf.close()
    return size

def compress_image_for_api(image: Image.Image, max_size_kb: int = 900) -> Image.Image:
    if image.mode in ('RGBA', 'P'):
        image = image.convert('RGB')
    if get_image_size_kb(image, 'JPEG', 85) <= max_size_kb:
        return image
    for q in [75, 65, 55, 45, 35, 25]:
        if get_image_size_kb(image, 'JPEG', q) <= max_size_kb:
            buf = BytesIO()
            image.save(buf, format='JPEG', quality=q, optimize=True)
            buf.seek(0)
            return Image.open(buf).convert('RGB')
    scale = 0.8
    while scale > 0.3:
        new_w = int(image.width * scale)
        new_h = int(image.height * scale)
        resized = image.resize((new_w, new_h), Image.LANCZOS)
        if get_image_size_kb(resized, 'JPEG', 65) <= max_size_kb:
            return resized
        scale -= 0.1
    max_dim = 1200
    if image.width > max_dim or image.height > max_dim:
        ratio = max_dim / max(image.width, image.height)
        new_w = int(image.width * ratio)
        new_h = int(image.height * ratio)
        image = image.resize((new_w, new_h), Image.LANCZOS)
    return image

def ocr_image_from_api(image: Image.Image) -> str:
    global request_count
    image = compress_image_for_api(image)
    for attempt in range(MAX_RETRIES):
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tmp:
                image.save(tmp.name, format="JPEG", quality=85, optimize=True)
                tmp_path = tmp.name
            with open(tmp_path, "rb") as f:
                files = {"file": f}
                data = {
                    "apikey": API_KEY,
                    "OCREngine": OCR_ENGINE,
                    "isOverlayRequired": False,
                }
                if LANGUAGE:
                    data["language"] = LANGUAGE
                response = requests.post(OCR_API_URL, files=files, data=data, timeout=60)
            if response.status_code != 200:
                print(f" HTTP {response.status_code}", end="")
                if attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_DELAY * (attempt + 1))
                    continue
                return ""
            try:
                result = response.json()
            except ValueError:
                print(" Non‑JSON response", end="")
                if attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_DELAY * (attempt + 1))
                    continue
                return ""
            if not isinstance(result, dict):
                print(f" Unexpected type: {type(result)}", end="")
                return ""
            if result.get("IsErroredOnProcessing"):
                error_msg = result.get("ErrorMessage", ["Unknown"])[0]
                print(f" OCR error: {error_msg[:100]}", end="")
                if "limit" in error_msg.lower() or "quota" in error_msg.lower():
                    print("\nQuota exceeded. Exiting.")
                    sys.exit(1)
                if attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_DELAY * (attempt + 1))
                    continue
                return ""
            request_count += 1
            return result["ParsedResults"][0]["ParsedText"]
        except Exception as e:
            print(f" Exception: {e}", end="")
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY * (attempt + 1))
        finally:
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.unlink(tmp_path)
                except PermissionError:
                    pass
    return ""

def find_title_and_group(text: str) -> Tuple[Optional[str], Optional[int]]:
    title_pattern = re.compile(r'Evidencijska lista\s+([0-9]+[a-z]?)\s*[-–]\s*([0-9]{1,4})', re.IGNORECASE)
    title_match = title_pattern.search(text)
    title_str = None
    if title_match:
        title_str = f"{title_match.group(1)}-{title_match.group(2)}"
    group_pattern = re.compile(r'Grupa\s*([12])', re.IGNORECASE)
    group_match = group_pattern.search(text)
    group_num = int(group_match.group(1)) if group_match else None
    return title_str, group_num

def load_excel_lookup(excel_path: str) -> Dict[Tuple[int, str, str], str]:
    lookup = {}
    for group, sheet_name in [(1, "G1_ planiranje"), (2, "G2_ planiranje")]:
        try:
            df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, dtype=str, engine='openpyxl')
            for _, row in df.iterrows():
                if len(row) < 5:
                    continue
                el_cell = str(row[2]) if pd.notna(row[2]) else ""
                id_cell = str(row[3]) if pd.notna(row[3]) else ""
                model_cell = str(row[4]) if pd.notna(row[4]) else ""
                if not el_cell.strip() or not id_cell.strip():
                    continue
                el_list = [el.strip() for el in re.split(r'[\n\r]+', el_cell) if el.strip()]
                for el in el_list:
                    lookup[(group, el, id_cell)] = model_cell
        except Exception as e:
            print(f"Warning: Could not load sheet '{sheet_name}': {e}")
    return lookup

def process_pdf(input_pdf, output_dir, rotate, dpi, lookup, number, all_rows, flagged_details, seen_combinations):
    """
    Process a single PDF file.
    flagged_details: list to append (filepath, list_of_reason_strings) – each reason is detailed.
    seen_combinations: dict {(group, el, id): first_file} to track cross‑file duplicates.
    Returns True if processed, False otherwise.
    """
    print(f"\n--- Processing: {input_pdf} ---")
    try:
        images = convert_from_path(input_pdf, dpi=dpi)
    except Exception as e:
        print(f"Failed to convert PDF: {e}")
        return False

    total_pages = len(images)
    if rotate:
        images = [img.rotate(180, expand=True) for img in images]

    markers = []   # (page_idx, title_str, group_num)
    for idx, img in enumerate(images):
        print(f"  Page {idx+1}/{total_pages}...", end="", flush=True)
        text = ocr_image_from_api(img)
        if not text:
            print(" [OCR failed]")
            continue
        title_str, group_num = find_title_and_group(text)
        if title_str and group_num is not None:
            markers.append((idx, title_str, group_num))
            print(f" [found: {title_str}, Grupa {group_num}]")
        else:
            print(" [no marker]")

    if not markers:
        print("No markers found in this PDF.")
        return False

    reasons = []

    # 1. Single page
    if total_pages == 1:
        reasons.append(f"File has only 1 page (contains {len(markers)} marker(s))")

    # 2. >2 pages but only 1 marker
    if total_pages > 2 and len(markers) == 1:
        reasons.append(f"File has {total_pages} pages but only 1 marker found")

    # 3. Internal duplicate (same EL-ID inside this file)
    seen_inside = {}
    internal_dup_found = False
    dup_inside = None
    for _, title_str, _ in markers:
        if title_str in seen_inside:
            internal_dup_found = True
            dup_inside = title_str
            break
        seen_inside[title_str] = True
    if internal_dup_found:
        reasons.append(f"Internal duplicate: EL-ID '{dup_inside}' appears twice in the same file")

    # 4. Cross‑file duplicates
    # Build list of combinations for this file
    file_combs = []
    for _, title_str, group_num in markers:
        el_part, id_part = (title_str.split('-', 1) if '-' in title_str else (title_str, ""))
        file_combs.append((group_num, el_part, id_part, title_str))

    cross_dup_details = []
    for comb in file_combs:
        group, el, id_, full_title = comb
        key = (group, el, id_)
        if key in seen_combinations:
            prev_file = seen_combinations[key]
            cross_dup_details.append(f"'{full_title}' (already in {os.path.basename(prev_file)})")
    if cross_dup_details:
        reasons.append(f"Cross‑file duplicate: {', '.join(cross_dup_details)}")

    # Add this file's combinations to the global tracker (after checking)
    for comb in file_combs:
        group, el, id_, full_title = comb
        seen_combinations[(group, el, id_)] = input_pdf   # store first occurrence

    if reasons:
        flagged_details.append((input_pdf, reasons))

    # Group pages and save split PDFs
    groups = []
    for i, (start_idx, title_str, group_num) in enumerate(markers):
        end_idx = markers[i+1][0] if i+1 < len(markers) else total_pages
        groups.append((start_idx, end_idx, title_str, group_num))

    reader = PdfReader(input_pdf)
    for start, end, title_str, group_num in groups:
        el_part, id_part = (title_str.split('-', 1) if '-' in title_str else (title_str, ""))
        key = (group_num, el_part, id_part)
        model = lookup.get(key, "")
        if not model:
            print(f"  Warning: No Excel match for G{group_num} EL {el_part} ID {id_part}")
        all_rows.append((el_part, id_part, model, number))

        writer = PdfWriter()
        for page_num in range(start, end):
            page = reader.pages[page_num]
            if rotate:
                page.rotate(180)
            writer.add_page(page)
        safe_title = title_str.replace('/', '_').replace(' ', '_')
        out_path = os.path.join(output_dir, f"G{group_num} ELO {safe_title}.pdf")
        with open(out_path, "wb") as f:
            writer.write(f)
        print(f"Saved: {out_path} (pages {start+1}-{end})")

    return True

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--group", required=True, help="Directory containing PDF files")
    parser.add_argument("--output-dir", default=".", help="Output directory")
    parser.add_argument("--rotate", action="store_true", help="Rotate pages 180°")
    parser.add_argument("--dpi", type=int, default=300, help="DPI for conversion")
    parser.add_argument("--number", type=str, required=True, help="Number to add to CSV")
    parser.add_argument("--append", action="store_true", help="Append to CSV")
    args = parser.parse_args()

    if not os.path.isdir(args.group):
        print(f"Error: Directory '{args.group}' not found.")
        sys.exit(1)

    os.makedirs(args.output_dir, exist_ok=True)

    print(f"Loading Excel data from {EXCEL_PATH}...")
    lookup = load_excel_lookup(EXCEL_PATH)
    print(f"Loaded {len(lookup)} mappings.")

    pdf_files = [os.path.join(args.group, f) for f in os.listdir(args.group) if f.lower().endswith('.pdf')]
    if not pdf_files:
        print(f"No PDF files in {args.group}")
        sys.exit(1)
    print(f"Found {len(pdf_files)} PDF files.\n")

    all_rows = []
    flagged_details = []   # (filepath, list_of_reason_strings)
    seen_combinations = {}  # (group, el, id) -> first file path

    total_requests_before = request_count
    for pdf_file in pdf_files:
        process_pdf(pdf_file, args.output_dir, args.rotate, args.dpi, lookup, args.number, all_rows, flagged_details, seen_combinations)

    # Write CSV
    csv_path = os.path.join(args.output_dir, "models.csv")
    mode = 'a' if args.append else 'w'
    with open(csv_path, mode, newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if not args.append:
            writer.writerow(['EL', 'ID', 'Model', 'Number'])
        writer.writerows(all_rows)

    print(f"\nTotal API requests: {request_count - total_requests_before}")
    print(f"CSV written to {csv_path} (appended={args.append})")

    # Detailed flagged report
    if flagged_details:
        print("\n" + "="*60)
        print("FLAGGED FILES (detailed):")
        for path, reasons in flagged_details:
            print(f"\n  📄 {os.path.basename(path)}")
            for r in reasons:
                print(f"     - {r}")
        print("="*60)
    else:
        print("\nNo files flagged.")

    print("Done.")

if __name__ == "__main__":
    main()