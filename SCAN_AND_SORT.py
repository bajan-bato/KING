#!/usr/bin/env python3
"""
Combined PDF splitter + sorter using OCR.space API.
Finds "Evidencijska lista XXX-YYYY" and "Grupa 1/2", groups pages,
extracts model name, and sorts into subfolders based on model name.
No Excel required – uses fuzzy matching against existing folder names.

Usage:
    python splitter_sorter.py input.pdf --rotate --output-dir output --sort-base sort
    python splitter_sorter.py --undo   # undo last sorting operation
"""

import os
import re
import sys
import csv
import time
import json
import shutil
import difflib
import argparse
import tempfile
from io import BytesIO
from typing import List, Tuple, Optional
from datetime import datetime

import requests
from PIL import Image
from pdf2image import convert_from_path
from PyPDF2 import PdfReader, PdfWriter

# ========== OCR CONFIGURATION ==========
API_KEY = "K84455376988957"
OCR_API_URL = "https://api.ocr.space/parse/image"
OCR_ENGINE = 2
LANGUAGE = ""
MAX_RETRIES = 3
RETRY_DELAY = 5

request_count = 0

# ========== SORTING CONFIGURATION ==========
SORT_BASE = "sort"          # can be overridden by --sort-base
OUTPUT_DIR = "output"       # temporary folder for split PDFs
LOG_FILE = "sort_log.txt"   # logs copy operations for undo

# ========== OCR FUNCTIONS (unchanged from your script) ==========
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

def find_title_and_group(text: str) -> Tuple[Optional[str], Optional[int], Optional[str]]:
    title_pattern = re.compile(r'Evidencijska lista\s+([0-9]+[a-z]?)\s*[-–]\s*([0-9]{1,4})', re.IGNORECASE)
    title_match = title_pattern.search(text)
    title_str = None
    if title_match:
        title_str = f"{title_match.group(1)}-{title_match.group(2)}"
    group_pattern = re.compile(r'Grupa\s*([12])', re.IGNORECASE)
    group_match = group_pattern.search(text)
    group_num = int(group_match.group(1)) if group_match else None
    model_name = ""
    model_match = re.search(r'Model\s*\n\s*([^\n]+)', text, re.IGNORECASE)
    if not model_match:
        model_match = re.search(r'Model\s*[:;]\s*([^\n]+)', text, re.IGNORECASE)
    if model_match:
        model_name = model_match.group(1).strip()
    return title_str, group_num, model_name

# ========== SORTING FUNCTIONS (no Excel) ==========
def find_target_folder(model_name: str, group: int, sort_base: str) -> Optional[str]:
    """Find best matching subfolder inside sort_base/G{group} using substring and fuzzy matching."""
    base_dir = os.path.join(sort_base, f"G{group}")
    if not os.path.isdir(base_dir):
        return None
    folders = [f for f in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, f))]
    if not folders:
        return None

    def normalize(s):
        s = s.lower()
        s = re.sub(r'^\d+\.\s*g[12]\s+', '', s)   # remove "1. G1 " prefix if any
        s = re.sub(r'[^\w\s]', ' ', s)
        s = re.sub(r'\s+', ' ', s).strip()
        return s

    norm_model = normalize(model_name)

    # 1. Substring match
    for folder in folders:
        norm_folder = normalize(folder)
        if norm_model in norm_folder or norm_folder in norm_model:
            return os.path.join(base_dir, folder)

    # 2. Fuzzy ratio fallback
    best_match = None
    best_score = -1
    for folder in folders:
        norm_folder = normalize(folder)
        score = difflib.SequenceMatcher(None, norm_model, norm_folder).ratio()
        if score > best_score:
            best_score = score
            best_match = folder
    if best_match:
        return os.path.join(base_dir, best_match)
    return None

def log_copy(operation_id: str, src: str, dst: str, model: str, group: int):
    """Append a record to the log file."""
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(f"{operation_id}|{src}|{dst}|{model}|{group}\n")

def undo_last_run():
    """Read the log file, find all entries from the most recent operation_id, delete those files."""
    if not os.path.exists(LOG_FILE):
        print("No log file found. Nothing to undo.")
        return

    with open(LOG_FILE, 'r', encoding='utf-8') as f:
        lines = [line.strip() for line in f if line.strip()]

    if not lines:
        print("Log file is empty. Nothing to undo.")
        return

    # Last line contains the last operation_id
    last_op_id = lines[-1].split('|')[0]
    # Collect all entries with that operation_id
    to_undo = []
    for line in lines:
        parts = line.split('|')
        if parts[0] == last_op_id:
            to_undo.append(parts)

    if not to_undo:
        print("No entries found for the last run.")
        return

    print(f"Found {len(to_undo)} files copied in the last run (ID: {last_op_id}).")
    confirm = input("Delete these files from their sorted folders? (y/N): ")
    if confirm.lower() != 'y':
        print("Aborted.")
        return

    for parts in to_undo:
        dst_path = parts[2]   # destination path
        if os.path.exists(dst_path):
            os.remove(dst_path)
            print(f"Removed: {dst_path}")
        else:
            print(f"Already missing: {dst_path}")

    # Optionally remove empty directories? Not necessary but can be added.
    # Remove the log entries of this run (or keep them but mark as undone)
    # We'll simply rewrite the log without the undone entries.
    remaining = [line for line in lines if line.split('|')[0] != last_op_id]
    with open(LOG_FILE, 'w', encoding='utf-8') as f:
        for line in remaining:
            f.write(line + '\n')
    print(f"Undo complete. Log updated.")

# ========== MAIN (combined) ==========
def main():
    parser = argparse.ArgumentParser(description="Split scanned PDF and sort by model name (no Excel).")
    parser.add_argument("input_pdf", nargs='?', help="Input PDF file (required unless --undo)")
    parser.add_argument("--rotate", action="store_true", help="Rotate pages 180°")
    parser.add_argument("--output-dir", default="output", help="Temp folder for split PDFs (default: output)")
    parser.add_argument("--sort-base", default="sort", help="Base folder for sorted subfolders (default: sort)")
    parser.add_argument("--dpi", type=int, default=300, help="DPI for conversion")
    parser.add_argument("--undo", action="store_true", help="Undo the last sorting operation (delete copied files)")
    args = parser.parse_args()

    if args.undo:
        undo_last_run()
        return

    if not args.input_pdf:
        print("Error: input PDF required unless --undo is used.")
        sys.exit(1)

    # Create folders
    os.makedirs(args.output_dir, exist_ok=True)
    os.makedirs(args.sort_base, exist_ok=True)
    for g in ["G1", "G2"]:
        os.makedirs(os.path.join(args.sort_base, g), exist_ok=True)

    # Generate a unique operation ID (timestamp)
    op_id = datetime.now().strftime("%Y%m%d_%H%M%S")

    # --- Step 1: OCR splitting (same as before) ---
    print(f"Converting PDF to images (dpi={args.dpi})...")
    try:
        images = convert_from_path(args.input_pdf, dpi=args.dpi)
    except Exception as e:
        print(f"Failed to convert PDF: {e}")
        sys.exit(1)

    total_pages = len(images)
    if args.rotate:
        print("Rotating all images...")
        images = [img.rotate(180, expand=True) for img in images]

    print("Running OCR via API...")
    markers = []          # (page_index, title_str, group_num)
    models_data = []      # for CSV

    for idx, img in enumerate(images):
        print(f"  Page {idx+1}/{total_pages}...", end="", flush=True)
        text = ocr_image_from_api(img)
        if not text:
            print(" [OCR failed]")
            continue
        title_str, group_num, model_name = find_title_and_group(text)
        if title_str and group_num is not None:
            markers.append((idx, title_str, group_num))
            el_part, id_part = (title_str.split('-', 1) if '-' in title_str else (title_str, ""))
            models_data.append({
                'EL': el_part,
                'ID': id_part,
                'Model': model_name
            })
            print(f" [found: {title_str}, Grupa {group_num}, Model: {model_name[:50]}]")
        else:
            print(" [no marker]")

    if not markers:
        print("No markers found. Exiting.")
        sys.exit(1)

    # Save CSV (optional)
    if models_data:
        csv_path = os.path.join(args.output_dir, "models.csv")
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=['EL', 'ID', 'Model'])
            writer.writeheader()
            writer.writerows(models_data)
        print(f"\nSaved model data to {csv_path}")

    # Group pages
    groups = []
    for i, (start_idx, title_str, group_num) in enumerate(markers):
        end_idx = markers[i+1][0] if i+1 < len(markers) else total_pages
        groups.append((start_idx, end_idx, title_str, group_num))

    # --- Step 2: Extract each group, save to output dir, then sort ---
    print("Extracting groups and sorting...")
    reader = PdfReader(args.input_pdf)
    for start, end, title_str, group_num in groups:
        # Extract PDF pages
        writer = PdfWriter()
        for page_num in range(start, end):
            page = reader.pages[page_num]
            if args.rotate:
                page.rotate(180)
            writer.add_page(page)
        safe_title = title_str.replace('/', '_').replace(' ', '_')
        filename = f"G{group_num} ELO {safe_title}.pdf"
        temp_path = os.path.join(args.output_dir, filename)
        with open(temp_path, "wb") as f:
            writer.write(f)
        print(f"Created: {temp_path}")

        # Find model name for this group (from models_data using title_str)
        model_name = ""
        for entry in models_data:
            if f"{entry['EL']}-{entry['ID']}" == title_str:
                model_name = entry['Model']
                break

        if not model_name:
            print(f"  WARNING: No model name extracted for {title_str}, skipping sort.")
            continue

        # Find target folder based on model name and group
        target_dir = find_target_folder(model_name, group_num, args.sort_base)
        if target_dir is None:
            print(f"  FAIL: No matching folder for model '{model_name}' (group {group_num})")
            continue

        # Copy file to target folder
        dst_path = os.path.join(target_dir, filename)
        if os.path.exists(dst_path):
            print(f"  SKIP: {filename} already exists in {target_dir}")
        else:
            shutil.copy2(temp_path, dst_path)
            print(f"  OK: {filename} -> {os.path.basename(target_dir)}")
            # Log the copy for undo
            log_copy(op_id, temp_path, dst_path, model_name, group_num)

    # Optional: delete temp output files? We'll leave them for now; user can clean up.
    print(f"\nTotal API requests used: {request_count}")
    print(f"Done. Sorted files are in '{args.sort_base}'. Use --undo to revert this run.")

if __name__ == "__main__":
    main()