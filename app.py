# app.py
import os
import uuid
import time
import threading
import shutil
from flask import Flask, request, jsonify, send_from_directory, abort

# --- Config (adjust if you want absolute paths) ---
ROOT = os.path.abspath(os.getcwd())               # project root (where app.py lives)
UPLOADS = os.path.join(ROOT, "uploads")
OUTPUTS = os.path.join(ROOT, "output")
TOOLS   = os.path.join(ROOT, "tools")

os.makedirs(UPLOADS, exist_ok=True)
os.makedirs(OUTPUTS, exist_ok=True)
os.makedirs(TOOLS, exist_ok=True)

# Tool default paths (will be used by convert function)
TESSERACT = os.environ.get("TESSERACT_PATH") or os.path.join(TOOLS, "tesseract", "tesseract.exe")
POPPLER_BIN = os.environ.get("POPPLER_PATH") or os.path.join(TOOLS, "poppler", "bin")
JAVA_EXE = os.environ.get("JAVA_PATH") or os.path.join(TOOLS, "java", "bin", "java.exe")
TABULA_JAR = os.environ.get("TABULA_JAR") or os.path.join(TOOLS, "tabula", "tabula.jar")

# Try to import the user's convert_all_in_one.convert_pdf_to_docx function if present
convert_func = None
try:
    import convert_all_in_one as converter_module
    if hasattr(converter_module, "convert_pdf_to_docx"):
        convert_func = converter_module.convert_pdf_to_docx
except Exception:
    convert_func = None

# If convert_all_in_one not present, we'll implement a minimal fallback using pdf2docx + OCR
def fallback_convert(pdf_path, out_docx):
    """
    Simple fallback conversion:
     - tries pdf2docx for digital PDFs
     - otherwise rasterizes pages via poppler (if available) and runs tesseract OCR
    This is intentionally conservative and intended as a working fallback.
    """
    from pdf2docx import Converter as PDF2DOCX
    from pdf2image import convert_from_path
    from docx import Document
    import pytesseract
    # set tesseract path if available
    if os.path.exists(TESSERACT):
        try:
            import pytesseract as _pt
            _pt.pytesseract.tesseract_cmd = TESSERACT
        except Exception:
            pass

    # quick embedded text detection (first 3 pages)
    has_text = False
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            sample = "".join((pdf.pages[i].extract_text() or "") for i in range(min(3, len(pdf.pages))))
            has_text = len(sample.strip()) > 50
    except Exception:
        has_text = False

    # If PDF has selectable text, try pdf2docx
    if has_text:
        try:
            cv = PDF2DOCX(pdf_path)
            cv.convert(out_docx, start=0, end=None)
            cv.close()
            return True
        except Exception:
            pass

    # fallback OCR per-page (high-DPI)
    imgs = convert_from_path(pdf_path, dpi=300, poppler_path=POPPLER_BIN if os.path.exists(POPPLER_BIN) else None)
    doc = Document()
    for i, img in enumerate(imgs):
        txt = pytesseract.image_to_string(img)
        for line in txt.splitlines():
            if line.strip():
                doc.add_paragraph(line)
        if i != len(imgs) - 1:
            doc.add_page_break()

    doc.save(out_docx)
    return True

# use selected conversion function (user-provided or fallback)
def run_conversion(pdf_path, out_docx, progress_callback=None):
    # prefer user convert_func if available
    if convert_func:
        # try calling user implementation; many user scripts accept (pdf_path, out_docx)
        try:
            # If convert_func uses tools relative to project root, ensure current env knows them
            os.environ["TESSERACT_PATH"] = TESSERACT
            os.environ["POPPLER_PATH"] = POPPLER_BIN
            os.environ["JAVA_PATH"] = JAVA_EXE
            os.environ["TABULA_JAR"] = TABULA_JAR
            return bool(convert_func(pdf_path, out_docx))
        except Exception as e:
            # fallback on error
            print("convert_all_in_one failed:", e)
            return fallback_convert(pdf_path, out_docx)
    else:
        return fallback_convert(pdf_path, out_docx)


# In-memory job store (for simplicity)
jobs = {}  # job_id -> {status, progress, in, out, error}

app = Flask(__name__, static_folder='.', static_url_path='')

@app.route("/")
def index():
    return app.send_static_file("index.html")

# helper
def allowed_pdf(fname):
    return fname.lower().endswith(".pdf")

@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return "no file", 400
    f = request.files["file"]
    if f.filename == "":
        return "empty filename", 400
    if not allowed_pdf(f.filename):
        return "only pdf allowed", 400

    job_id = str(uuid.uuid4())
    in_name = job_id + ".pdf"
    in_path = os.path.join(UPLOADS, in_name)
    f.save(in_path)

    out_name = job_id + ".docx"
    out_path = os.path.join(OUTPUTS, out_name)

    jobs[job_id] = {"status": "queued", "progress": 0, "in": in_name, "out": None, "error": None}

    # start background thread
    thread = threading.Thread(target=_worker, args=(job_id, in_path, out_path), daemon=True)
    thread.start()

    return jsonify({"id": job_id})

def _worker(job_id, in_path, out_path):
    try:
        jobs[job_id]["status"] = "processing"
        jobs[job_id]["progress"] = 5

        # run conversion (report rough progress)
        jobs[job_id]["progress"] = 10
        success = run_conversion(in_path, out_path, progress_callback=None)

        # ensure file was created and is reasonable size
        if not success or not os.path.exists(out_path) or os.path.getsize(out_path) < 1024:
            # Try a final fallback: if pdf2docx can at least export something
            try:
                # attempt one more time with fallback converter
                ok = fallback_convert(in_path, out_path)
            except Exception as e:
                ok = False
                print("fallback final attempt failed:", e)

            if (not ok) or (not os.path.exists(out_path)) or os.path.getsize(out_path) < 1024:
                jobs[job_id]["status"] = "error"
                jobs[job_id]["error"] = "Conversion failed or output corrupted (file missing or too small)."
                jobs[job_id]["progress"] = 0
                return

        # success
        jobs[job_id]["out"] = os.path.basename(out_path)
        jobs[job_id]["progress"] = 100
        jobs[job_id]["status"] = "done"
    except Exception as e:
        jobs[job_id]["status"] = "error"
        jobs[job_id]["error"] = str(e)
        jobs[job_id]["progress"] = 0

@app.route("/status")
def status():
    job_id = request.args.get("id")
    if not job_id or job_id not in jobs:
        return jsonify({"status": "error", "message": "invalid id"}), 404
    j = jobs[job_id]
    return jsonify({
        "status": j.get("status"),
        "progress": j.get("progress", 0),
        "out": j.get("out"),
        "message": j.get("error")
    })

@app.route("/download")
def download():
    fname = request.args.get("file")
    if not fname:
        return "missing file", 400
    safe = os.path.basename(fname)
    path = os.path.join(OUTPUTS, safe)
    if not os.path.exists(path):
        return "file not found", 404
    # final validation: docx must be > 1 KB
    if os.path.getsize(path) < 1024:
        return "file corrupted or too small", 500
    return send_from_directory(OUTPUTS, safe, as_attachment=True)

if __name__ == "__main__":
    # helpful startup messages
    print("Project root:", ROOT)
    print("Uploads folder:", UPLOADS)
    print("Outputs folder:", OUTPUTS)
    print("Tools lookup:", TOOLS)
    print("Tesseract path:", TESSERACT)
    print("Poppler bin:", POPPLER_BIN)
    print("Java exe:", JAVA_EXE)
    print("Tabula jar:", TABULA_JAR)
    app.run(host="0.0.0.0", port=5000, debug=True)
