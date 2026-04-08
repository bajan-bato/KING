#!/usr/bin/env python3
"""
Split scanned PDF into groups using OCR.space API.
Finds "Evidencijska lista XXX-YYYY" and "Grupa 1/2", groups pages accordingly.

Usage:
python3 PDFscanner.py input.pdf --rotate --output-dir ./output
"""

import os
import re
import sys
import time
import argparse
import tempfile
from typing import List, Tuple, Optional

import requests
from PIL import Image
from pdf2image import convert_from_path
from PyPDF2 import PdfReader, PdfWriter

# ----------------------------------------------------------------------
# OCR.space API settings (your key inserted)
API_KEY = "K84455376988957"
OCR_API_URL = "https://api.ocr.space/parse/image"
OCR_ENGINE = 2
LANGUAGE = ""                       # Croatian
MAX_RETRIES = 3
RETRY_DELAY = 5                        # seconds between retries
# ----------------------------------------------------------------------

def compress_image_for_api(image: Image.Image, max_size_kb: int = 900) -> Image.Image:
    """Compress image to stay under API file size limit (1024 KB)."""
    # Check current size
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        image.save(tmp.name, format="PNG")
        size_kb = os.path.getsize(tmp.name) / 1024
        os.unlink(tmp.name)
        if size_kb <= max_size_kb:
            return image

    # Scale down
    scale = 0.9
    resample_filter = getattr(Image, 'LANCZOS', Image.BICUBIC)
    while scale > 0.5:
        new_size = (int(image.width * scale), int(image.height * scale))
        resized = image.resize(new_size, resample_filter)
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            resized.save(tmp.name, format="PNG", optimize=True)
            size_kb = os.path.getsize(tmp.name) / 1024
            os.unlink(tmp.name)
            if size_kb <= max_size_kb:
                return resized
        scale -= 0.1

    # Last resort: JPEG with lower quality
    with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tmp:
        image.save(tmp.name, format="JPEG", quality=50)
        size_kb = os.path.getsize(tmp.name) / 1024
        os.unlink(tmp.name)
        if size_kb <= max_size_kb:
            return Image.open(tmp.name).convert("RGB")
    return image

def ocr_image_from_api(image: Image.Image) -> str:
    """Send image to OCR.space API with retries and robust error handling."""
    image = compress_image_for_api(image)

    for attempt in range(MAX_RETRIES):
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                image.save(tmp.name, format="PNG", optimize=True)
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

            # Check HTTP status
            if response.status_code != 200:
                print(f" HTTP {response.status_code}", end="")
                if attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_DELAY * (attempt + 1))
                    continue
                return ""

            # Try to parse JSON
            try:
                result = response.json()
            except ValueError:
                print(" Non‑JSON response", end="")
                if attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_DELAY * (attempt + 1))
                    continue
                return ""

            # Check for error in result
            if not isinstance(result, dict):
                print(f" Unexpected type: {type(result)}", end="")
                return ""

            if result.get("IsErroredOnProcessing"):
                error_msg = result.get("ErrorMessage", ["Unknown"])[0]
                print(f" OCR error: {error_msg[:100]}", end="")
                # Quota exceeded – stop immediately
                if "limit" in error_msg.lower() or "quota" in error_msg.lower():
                    print("\nQuota exceeded. Exiting.")
                    sys.exit(1)
                if attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_DELAY * (attempt + 1))
                    continue
                return ""

            # Success
            return result["ParsedResults"][0]["ParsedText"]

        except Exception as e:
            print(f" Exception: {e}", end="")
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY * (attempt + 1))
        finally:
            if tmp_path and os.path.exists(tmp_path):
                os.unlink(tmp_path)

    return ""

def find_title_and_group(text: str) -> Tuple[Optional[str], Optional[int]]:
    title_pattern = re.compile(
        r'Evidencijska lista\s+([0-9]+[a-z]?)\s*[-–]\s*([0-9]{1,4})',
        re.IGNORECASE
    )
    title_match = title_pattern.search(text)
    title_str = None
    if title_match:
        num_part = title_match.group(1)
        year_part = title_match.group(2)
        title_str = f"{num_part}-{year_part}"

    group_pattern = re.compile(r'Grupa\s*([12])', re.IGNORECASE)
    group_match = group_pattern.search(text)
    group_num = None
    if group_match:
        group_num = int(group_match.group(1))

    return title_str, group_num

def main():
    parser = argparse.ArgumentParser(
        description="Split scanned PDF by 'Evidencijska lista' headers using OCR.space API."
    )
    parser.add_argument("input_pdf", help="Path to the input PDF file")
    parser.add_argument("--rotate", action="store_true",
                        help="Rotate every page by 180 degrees before processing")
    parser.add_argument("--output-dir", default=".",
                        help="Directory to save output PDFs (default: current directory)")
    parser.add_argument("--dpi", type=int, default=300,
                        help="DPI for PDF->image conversion (default: 300)")
    args = parser.parse_args()

    if not os.path.isfile(args.input_pdf):
        print(f"Input file not found: {args.input_pdf}")
        sys.exit(1)

    os.makedirs(args.output_dir, exist_ok=True)

    # Convert PDF to images
    print(f"Converting PDF to images (dpi={args.dpi})...")
    try:
        images = convert_from_path(args.input_pdf, dpi=args.dpi)
    except Exception as e:
        print("Failed to convert PDF to images. Is poppler-utils installed?")
        print(f"Error: {e}")
        sys.exit(1)

    total_pages = len(images)
    if args.rotate:
        print("Rotating all images by 180 degrees...")
        images = [img.rotate(180, expand=True) for img in images]

    # OCR and find markers
    print("Running OCR via API (this may be slow)...")
    markers = []
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
        print("No markers found. Exiting.")
        sys.exit(1)

    # Build groups
    groups = []
    for i, (start_idx, title_str, group_num) in enumerate(markers):
        end_idx = markers[i+1][0] if i+1 < len(markers) else total_pages
        groups.append((start_idx, end_idx, title_str, group_num))

    # Extract and save groups
    print("Extracting groups from original PDF...")
    reader = PdfReader(args.input_pdf)
    for start, end, title_str, group_num in groups:
        writer = PdfWriter()
        for page_num in range(start, end):
            page = reader.pages[page_num]
            if args.rotate:
                page.rotate(180)
            writer.add_page(page)
        safe_title = title_str.replace('/', '_').replace(' ', '_')
        out_path = os.path.join(args.output_dir, f"G{group_num} ELO {safe_title}.pdf")
        with open(out_path, "wb") as f:
            writer.write(f)
        print(f"Saved: {out_path} (pages {start+1}-{end})")

    print("Done.")

if __name__ == "__main__":
    main()