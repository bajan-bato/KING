#!/usr/bin/env python3
"""
PDF Rotator Gallery – manually select files to rotate, then export.
Usage: python pdf_rotator_gallery.py
"""

import os
import base64
import webbrowser
import threading
from pathlib import Path
from io import BytesIO
from flask import Flask, render_template_string, request, jsonify
from pdf2image import convert_from_path
from PIL import Image
from PyPDF2 import PdfReader, PdfWriter

# ========== CONFIGURATION ==========
FOLDER = "rotate"        # folder containing original PDFs
OUTPUT_FOLDER = "rotated" # folder for rotated copies
PORT = 5000
# ===================================

app = Flask(__name__)

def get_thumbnail_base64(pdf_path, size=(200, 200)):
    """Return base64 JPEG thumbnail of first page."""
    try:
        images = convert_from_path(pdf_path, dpi=72, first_page=1, last_page=1)
        if not images:
            return None
        img = images[0]
        img.thumbnail(size, Image.LANCZOS)
        buffered = BytesIO()
        img.save(buffered, format="JPEG")
        return base64.b64encode(buffered.getvalue()).decode()
    except Exception as e:
        print(f"Error making thumbnail for {pdf_path}: {e}")
        return None

def rotate_pdf_file(src, dst):
    """Rotate all pages of a PDF by 180° and save to dst."""
    reader = PdfReader(src)
    writer = PdfWriter()
    for page in reader.pages:
        page.rotate(180)
        writer.add_page(page)
    with open(dst, "wb") as f:
        writer.write(f)

@app.route('/')
def gallery():
    """Serve the HTML gallery."""
    folder = Path(FOLDER)
    if not folder.is_dir():
        return f"Error: Folder '{FOLDER}' does not exist.", 500

    pdf_files = sorted(folder.glob("*.pdf"))
    if not pdf_files:
        return f"No PDF files found in '{FOLDER}'.", 500

    cards = []
    for pdf in pdf_files:
        thumb = get_thumbnail_base64(pdf)
        if thumb:
            cards.append({
                'name': pdf.name,
                'thumb': thumb
            })
        else:
            cards.append({
                'name': pdf.name,
                'thumb': None
            })

    html_template = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>PDF Rotator Gallery</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
            h1 { color: #333; }
            .toolbar { margin-bottom: 20px; }
            button { padding: 8px 16px; margin-right: 10px; font-size: 14px; cursor: pointer; background: #e0e0e0; border: 1px solid #aaa; border-radius: 4px; }
            button.primary { background: #4CAF50; color: white; border: none; }
            button.danger { background: #f44336; color: white; border: none; }
            .gallery { display: flex; flex-wrap: wrap; gap: 20px; }
            .card { width: 220px; background: white; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); padding: 10px; text-align: center; cursor: pointer; transition: transform 0.1s; }
            .card.selected { background: #ffebee; border: 2px solid #f44336; }
            .card img { max-width: 200px; max-height: 200px; border: 1px solid #ddd; border-radius: 4px; }
            .filename { font-size: 12px; margin-top: 8px; word-break: break-all; color: #555; }
            .status { margin-top: 20px; font-style: italic; }
        </style>
    </head>
    <body>
        <h1>PDF Rotator Gallery</h1>
        <p>Click on a PDF to mark it for rotation (red border). Use the buttons below to rotate all or export.</p>
        <div class="toolbar">
            <button id="rotateAllBtn" class="primary">Rotate all</button>
            <button id="clearAllBtn">Clear all</button>
            <button id="exportBtn" class="danger">Export rotated</button>
        </div>
        <div class="gallery" id="gallery">
            {% for card in cards %}
            <div class="card" data-filename="{{ card.name }}">
                {% if card.thumb %}
                <img src="data:image/jpeg;base64,{{ card.thumb }}" alt="{{ card.name }}">
                {% else %}
                <div style="width:200px; height:200px; background:#eee; display:flex; align-items:center; justify-content:center;">No preview</div>
                {% endif %}
                <div class="filename">{{ card.name }}</div>
            </div>
            {% endfor %}
        </div>
        <div class="status" id="status"></div>
        <script>
            let selected = new Set();
            const cards = document.querySelectorAll('.card');
            const statusDiv = document.getElementById('status');

            function updateUI() {
                cards.forEach(card => {
                    const fname = card.getAttribute('data-filename');
                    if (selected.has(fname)) {
                        card.classList.add('selected');
                    } else {
                        card.classList.remove('selected');
                    }
                });
                statusDiv.innerText = `Selected: ${selected.size} file(s)`;
            }

            cards.forEach(card => {
                card.addEventListener('click', (e) => {
                    const fname = card.getAttribute('data-filename');
                    if (selected.has(fname)) {
                        selected.delete(fname);
                    } else {
                        selected.add(fname);
                    }
                    updateUI();
                });
            });

            document.getElementById('rotateAllBtn').addEventListener('click', () => {
                cards.forEach(card => {
                    const fname = card.getAttribute('data-filename');
                    selected.add(fname);
                });
                updateUI();
            });

            document.getElementById('clearAllBtn').addEventListener('click', () => {
                selected.clear();
                updateUI();
            });

            document.getElementById('exportBtn').addEventListener('click', async () => {
                if (selected.size === 0) {
                    alert('No files selected.');
                    return;
                }
                const fileList = Array.from(selected);
                statusDiv.innerText = 'Exporting... please wait.';
                try {
                    const response = await fetch('/export', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ files: fileList })
                    });
                    const result = await response.json();
                    if (result.success) {
                        statusDiv.innerText = `Exported ${result.count} files to '{{ output_folder }}' folder.`;
                        alert(`Success! Rotated ${result.count} files saved to '{{ output_folder }}'.`);
                    } else {
                        statusDiv.innerText = 'Export failed: ' + result.error;
                        alert('Export failed: ' + result.error);
                    }
                } catch (err) {
                    statusDiv.innerText = 'Network error.';
                    alert('Network error: ' + err);
                }
            });
        </script>
    </body>
    </html>
    """
    return render_template_string(html_template, cards=cards, output_folder=OUTPUT_FOLDER)

@app.route('/export', methods=['POST'])
def export():
    """Rotate selected files and save to OUTPUT_FOLDER."""
    data = request.get_json()
    files_to_rotate = data.get('files', [])
    if not files_to_rotate:
        return jsonify({'success': False, 'error': 'No files selected'}), 400

    input_dir = Path(FOLDER)
    output_dir = Path(OUTPUT_FOLDER)
    output_dir.mkdir(parents=True, exist_ok=True)

    rotated_count = 0
    errors = []
    for fname in files_to_rotate:
        src = input_dir / fname
        if not src.is_file():
            errors.append(f"{fname} not found")
            continue
        dst = output_dir / fname
        try:
            rotate_pdf_file(src, dst)
            rotated_count += 1
        except Exception as e:
            errors.append(f"{fname}: {str(e)}")

    if errors:
        return jsonify({'success': False, 'error': '; '.join(errors)}), 500
    return jsonify({'success': True, 'count': rotated_count})

def open_browser():
    webbrowser.open(f'http://localhost:{PORT}')

if __name__ == '__main__':
    # Check if folder exists
    if not os.path.isdir(FOLDER):
        print(f"Error: Folder '{FOLDER}' does not exist. Create it and put PDFs inside.")
        exit(1)
    # Start browser after a short delay
    threading.Timer(1.5, open_browser).start()
    app.run(host='127.0.0.1', port=PORT, debug=False)