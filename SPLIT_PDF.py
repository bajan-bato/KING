#!/usr/bin/env python3
import sys
from pathlib import Path
from pypdf import PdfReader, PdfWriter

pdf_path = Path(sys.argv[1] if len(sys.argv) > 1 else input("PDF put: "))
split_page = int(input("Stranica reza (1-indeksirano): ")) - 1
name1 = input("Ime prvog dijela: ") or f"{pdf_path.stem}_dio1.pdf"
name2 = input("Ime drugog dijela: ") or f"{pdf_path.stem}_dio2.pdf"

reader = PdfReader(pdf_path)
writer1, writer2 = PdfWriter(), PdfWriter()
for i, page in enumerate(reader.pages):
    (writer1 if i < split_page else writer2).add_page(page)

out_dir = pdf_path.parent
writer1.write(out_dir / name1)
writer2.write(out_dir / name2)
print(f"Spremljeno: {name1} i {name2}")