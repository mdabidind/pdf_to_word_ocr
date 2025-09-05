import os
from pdf2docx import Converter

INPUT_DIR = "input"
OUTPUT_DIR = "output"

os.makedirs(OUTPUT_DIR, exist_ok=True)

for filename in os.listdir(INPUT_DIR):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(INPUT_DIR, filename)
        docx_name = os.path.splitext(filename)[0] + ".docx"
        docx_path = os.path.join(OUTPUT_DIR, docx_name)

        print(f"Converting {pdf_path} → {docx_path}")
        try:
            cv = Converter(pdf_path)
            cv.convert(docx_path, start=0, end=None)
            cv.close()
            print(f"✅ Saved {docx_path}")
        except Exception as e:
            print(f"❌ Failed to convert {filename}: {e}")
