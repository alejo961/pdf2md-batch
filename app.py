#!/usr/bin/env python3
"""
PDF2MD Batch Converter
=====================
Local web application for batch converting PDF and Word files to Markdown.
Uses pymupdf4llm for high-quality extraction with table, image, and header support.

Usage:
    python app.py
    Then open http://localhost:5000 in your browser.
"""

import os
import sys
import json
import time
import zipfile
import shutil
import threading
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor

from flask import (
    Flask, render_template_string, request, jsonify,
    send_file, send_from_directory
)

import pymupdf4llm
import fitz

try:
    from docx import Document
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except ImportError:
    Document = None
    CT_Tbl = None
    CT_P = None
    Table = None
    Paragraph = None

# ─── Configuration ───────────────────────────────────────────────────────────

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500 MB max total upload

UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("output")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

WORD_EXTENSIONS = {".docx", ".docm"}
SUPPORTED_CONVERT_EXTENSIONS = {".pdf", *WORD_EXTENSIONS}

# Track conversion jobs
jobs = {}
jobs_lock = threading.Lock()


# ─── Utility Logic ───────────────────────────────────────────────────────────

def is_pdf_scanned(pdf_path: Path) -> bool:
    """Check if the PDF is likely scanned (contains no text)."""
    try:
        doc = fitz.open(str(pdf_path))
        # Check first few pages
        for i in range(min(5, len(doc))):
            page = doc[i]
            # If page is searchable or has significant text, it's not scanned
            if hasattr(page, "is_searchable"):
                if page.is_searchable():
                    return False

            text = page.get_text().strip()
            if len(text) > 100:  # Threshold for "real" text
                return False
        return True
    except Exception as e:
        print(f"Error detecting scanned status for {pdf_path}: {e}")
        return True  # Default to true/scanned if error
    finally:
        if 'doc' in locals():
            doc.close()


def is_supported_convert_file(filename: str) -> bool:
    """Return True when the file can be converted to Markdown."""
    return Path(filename).suffix.lower() in SUPPORTED_CONVERT_EXTENSIONS


def markdown_escape_table_cell(value: str) -> str:
    """Escape Markdown table separators inside table cells."""
    return value.replace("|", "\\|").replace("\n", "<br>")


def iter_docx_blocks(document):
    """Yield Word paragraphs and tables in document order."""
    for child in document.element.body.iterchildren():
        if CT_P is not None and isinstance(child, CT_P):
            yield Paragraph(child, document)
        elif CT_Tbl is not None and isinstance(child, CT_Tbl):
            yield Table(child, document)


def docx_paragraph_to_markdown(paragraph) -> str:
    """Convert a Word paragraph to a Markdown line."""
    text_parts = []
    for run in paragraph.runs:
        run_text = run.text.replace("\t", "    ")
        if not run_text:
            continue
        if run.bold and run.italic:
            run_text = f"***{run_text}***"
        elif run.bold:
            run_text = f"**{run_text}**"
        elif run.italic:
            run_text = f"*{run_text}*"
        text_parts.append(run_text)

    text = "".join(text_parts).strip()
    if not text:
        text = paragraph.text.strip()
    if not text:
        return ""

    style_name = (paragraph.style.name if paragraph.style else "").lower()
    if style_name.startswith("heading"):
        digits = "".join(ch for ch in style_name if ch.isdigit())
        level = min(max(int(digits or "1"), 1), 6)
        return f"{'#' * level} {text}"
    if "list bullet" in style_name:
        return f"- {text}"
    if "list number" in style_name:
        return f"1. {text}"
    if style_name in {"title", "subtitle"}:
        return f"# {text}" if style_name == "title" else f"## {text}"
    return text


def docx_table_to_markdown(table) -> str:
    """Convert a Word table to GitHub-flavored Markdown."""
    rows = []
    for row in table.rows:
        cells = []
        for cell in row.cells:
            cell_text = "\n".join(
                docx_paragraph_to_markdown(paragraph)
                for paragraph in cell.paragraphs
                if paragraph.text.strip()
            )
            cells.append(markdown_escape_table_cell(cell_text.strip()))
        rows.append(cells)

    if not rows:
        return ""

    column_count = max(len(row) for row in rows)
    normalized_rows = [row + [""] * (column_count - len(row)) for row in rows]
    header = normalized_rows[0]
    separator = ["---"] * column_count
    body = normalized_rows[1:]

    md_rows = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(separator) + " |",
    ]
    md_rows.extend("| " + " | ".join(row) + " |" for row in body)
    return "\n".join(md_rows)


# ─── Conversion Logic ────────────────────────────────────────────────────────

def convert_pdf_to_md(pdf_path: Path, output_path: Path, job_id: str, file_index: int, options: dict = None):
    """Convert a single PDF file to Markdown using pymupdf4llm."""
    options = options or {}
    start_time = time.time()

    ocr_mode = options.get("ocr_mode", "auto")
    use_ocr = False

    # Determine if we need OCR
    if ocr_mode == "force":
        use_ocr = True
    elif ocr_mode == "auto":
        use_ocr = is_pdf_scanned(pdf_path)

    # Choose DPI: higher DPI improves OCR quality but increases processing time
    dpi = int(options.get("ocr_dpi") or (300 if use_ocr else 150))
    preprocess = options.get("ocr_preprocess", False)

    try:
        md_text = pymupdf4llm.to_markdown(
            str(pdf_path),
            write_images=options.get("extract_images", True),
            image_path=str(output_path.parent / "images"),
            image_format="png",
            dpi=dpi,
            force_ocr=use_ocr
        )

        output_path.write_bytes(md_text.encode("utf-8"))
        elapsed = round(time.time() - start_time, 2)

        with jobs_lock:
            jobs[job_id]["files"][file_index]["status"] = "done"
            jobs[job_id]["files"][file_index]["output"] = str(output_path.name)
            jobs[job_id]["files"][file_index]["size_md"] = len(md_text)
            jobs[job_id]["files"][file_index]["time"] = elapsed
            jobs[job_id]["completed"] += 1

    except Exception as e:
        with jobs_lock:
            jobs[job_id]["files"][file_index]["status"] = "error"
            jobs[job_id]["files"][file_index]["error"] = str(e)
            jobs[job_id]["completed"] += 1


def convert_word_to_md(word_path: Path, output_path: Path, job_id: str, file_index: int):
    """Convert a Word .docx/.docm file to Markdown."""
    start_time = time.time()

    try:
        if Document is None:
            raise RuntimeError("python-docx is not installed. Run: pip install python-docx")

        document = Document(str(word_path))
        md_blocks = []

        for block in iter_docx_blocks(document):
            if Paragraph is not None and isinstance(block, Paragraph):
                md_line = docx_paragraph_to_markdown(block)
            elif Table is not None and isinstance(block, Table):
                md_line = docx_table_to_markdown(block)
            else:
                md_line = ""

            if md_line:
                md_blocks.append(md_line)

        md_text = "\n\n".join(md_blocks).strip()
        if md_text:
            md_text += "\n"

        output_path.write_text(md_text, encoding="utf-8")
        elapsed = round(time.time() - start_time, 2)

        with jobs_lock:
            jobs[job_id]["files"][file_index]["status"] = "done"
            jobs[job_id]["files"][file_index]["output"] = str(output_path.name)
            jobs[job_id]["files"][file_index]["size_md"] = len(md_text)
            jobs[job_id]["files"][file_index]["time"] = elapsed
            jobs[job_id]["completed"] += 1

    except Exception as e:
        with jobs_lock:
            jobs[job_id]["files"][file_index]["status"] = "error"
            jobs[job_id]["files"][file_index]["error"] = str(e)
            jobs[job_id]["completed"] += 1


def convert_file_to_md(input_path: Path, output_path: Path, job_id: str, file_index: int, options: dict = None):
    """Convert one supported input file to Markdown based on its extension."""
    suffix = input_path.suffix.lower()
    if suffix == ".pdf":
        convert_pdf_to_md(input_path, output_path, job_id, file_index, options)
    elif suffix in WORD_EXTENSIONS:
        convert_word_to_md(input_path, output_path, job_id, file_index)
    else:
        with jobs_lock:
            jobs[job_id]["files"][file_index]["status"] = "error"
            jobs[job_id]["files"][file_index]["error"] = f"Unsupported file type: {suffix}"
            jobs[job_id]["completed"] += 1


def combine_md_files(job_output_dir: Path) -> Path:
    """Combine all MD files in a directory into a single file."""
    combined_path = job_output_dir / "all_combined.md"
    md_files = sorted([f for f in job_output_dir.iterdir() if f.suffix == ".md" and f.name != "all_combined.md"])

    with open(combined_path, "w", encoding="utf-8") as outf:
        for i, md_file in enumerate(md_files):
            if i > 0:
                outf.write("\n\n" + "=" * 80 + "\n\n")
            content = md_file.read_text(encoding="utf-8")
            outf.write(content)

    return combined_path


def merge_md_files(md_files_data: list, job_id: str) -> Path:
    """Merge uploaded MD files into a single file."""
    job_output_dir = OUTPUT_DIR / job_id
    job_output_dir.mkdir(parents=True, exist_ok=True)

    # Save individual files and create merged version
    merged_path = job_output_dir / "merged_markdown.md"

    with open(merged_path, "w", encoding="utf-8") as outf:
        for i, file_data in enumerate(md_files_data):
            if i > 0:
                outf.write("\n\n" + "=" * 80 + "\n\n")
            outf.write(file_data["content"])

    return merged_path


def run_batch_conversion(job_id: str):
    """Process all supported files in a job in parallel."""
    job = jobs[job_id]
    job_upload_dir = UPLOAD_DIR / job_id
    job_output_dir = OUTPUT_DIR / job_id
    job_output_dir.mkdir(parents=True, exist_ok=True)
    (job_output_dir / "images").mkdir(exist_ok=True)

    options = job.get("options", {})

    # Use ThreadPoolExecutor for parallel conversion
    max_workers = min(os.cpu_count() or 4, 8)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = []
        for i, file_info in enumerate(job["files"]):
            input_path = job_upload_dir / file_info["original_name"]
            md_name = Path(file_info["original_name"]).stem + ".md"
            output_path = job_output_dir / md_name

            with jobs_lock:
                file_info["status"] = "converting"

            futures.append(executor.submit(convert_file_to_md, input_path, output_path, job_id, i, options))

        # Wait for all conversions to finish
        for future in futures:
            try:
                future.result()
            except Exception as e:
                print(f"Error in future result: {e}")

    with jobs_lock:
        job["status"] = "done"
        job["finished_at"] = datetime.now().isoformat()

    # Create combined MD file
    combine_md_files(job_output_dir)

    # Create ZIP of all outputs
    zip_path = job_output_dir / "all_markdown.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in job_output_dir.iterdir():
            if f.suffix == ".md":
                zf.write(f, f.name)
        images_dir = job_output_dir / "images"
        if images_dir.exists():
            for img in images_dir.iterdir():
                zf.write(img, f"images/{img.name}")


# ─── Routes ──────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route("/upload", methods=["POST"])
def upload():
    """Handle batch PDF/Word upload and start conversion."""
    files = request.files.getlist("files") or request.files.getlist("pdfs")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "No files uploaded"}), 400

    job_id = datetime.now().strftime("%Y%m%d_%H%M%S") + f"_{id(files) % 10000:04d}"
    job_upload_dir = UPLOAD_DIR / job_id
    job_upload_dir.mkdir(parents=True, exist_ok=True)

    file_list = []
    for f in files:
        if f.filename and is_supported_convert_file(f.filename):
            safe_name = f.filename.replace("/", "_").replace("\\", "_")
            save_path = job_upload_dir / safe_name
            f.save(str(save_path))
            file_list.append({
                "original_name": safe_name,
                "size": save_path.stat().st_size,
                "status": "queued",
                "output": None,
                "error": None,
                "time": None,
                "size_md": None,
                "type": save_path.suffix.lower().lstrip("."),
            })

    if not file_list:
        return jsonify({"error": "No valid PDF or Word files found (.pdf, .docx, .docm)"}), 400

    # Options from request
    # Parse OCR DPI safely
    ocr_dpi_raw = request.form.get("ocrDpi")
    try:
        ocr_dpi = int(ocr_dpi_raw) if ocr_dpi_raw else None
    except Exception:
        ocr_dpi = None

    options = {
        "extract_images": request.form.get("optImages") == "true",
        "exclude_headers": request.form.get("optHeaders") == "true",
        "ocr_mode": request.form.get("ocrMode", "auto"),
        "ocr_dpi": ocr_dpi,
        "ocr_preprocess": request.form.get("ocrPreprocess") == "true"
    }

    with jobs_lock:
        jobs[job_id] = {
            "id": job_id,
            "status": "running",
            "created_at": datetime.now().isoformat(),
            "finished_at": None,
            "total": len(file_list),
            "completed": 0,
            "files": file_list,
            "options": options
        }

    # Run conversion in background thread
    thread = threading.Thread(target=run_batch_conversion, args=(job_id,), daemon=True)
    thread.start()

    return jsonify({"job_id": job_id, "total": len(file_list)})


@app.route("/merge-md", methods=["POST"])
def merge_md():
    """Handle MD files upload and merge them."""
    files = request.files.getlist("mds")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "No files uploaded"}), 400

    job_id = datetime.now().strftime("%Y%m%d_%H%M%S") + f"_merge_{id(files) % 10000:04d}"

    md_files_data = []
    for f in files:
        if f.filename and f.filename.lower().endswith(".md"):
            try:
                content = f.read().decode("utf-8")
                md_files_data.append({
                    "name": f.filename,
                    "content": content
                })
            except Exception as e:
                return jsonify({"error": f"Error reading {f.filename}: {str(e)}"}), 400

    if not md_files_data:
        return jsonify({"error": "No valid MD files found"}), 400

    # Merge files
    merge_md_files(md_files_data, job_id)

    with jobs_lock:
        jobs[job_id] = {
            "id": job_id,
            "status": "done",
            "type": "merge",
            "created_at": datetime.now().isoformat(),
            "finished_at": datetime.now().isoformat(),
            "total": len(md_files_data),
            "completed": len(md_files_data),
            "files": [{"original_name": f["name"], "status": "done"} for f in md_files_data],
        }

    return jsonify({"job_id": job_id, "total": len(md_files_data), "status": "done"})


@app.route("/status/<job_id>")
def status(job_id):
    """Get job status."""
    with jobs_lock:
        job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify(job)


@app.route("/download/<job_id>/<filename>")
def download_file(job_id, filename):
    """Download a single converted file."""
    job_output_dir = OUTPUT_DIR / job_id
    file_path = job_output_dir / filename
    if not file_path.exists():
        return jsonify({"error": "File not found"}), 404
    return send_file(str(file_path.resolve()), as_attachment=True)


@app.route("/download-all/<job_id>")
def download_all(job_id):
    """Download all converted files as ZIP."""
    zip_path = OUTPUT_DIR / job_id / "all_markdown.zip"
    if not zip_path.exists():
        return jsonify({"error": "ZIP not ready yet"}), 404
    return send_file(str(zip_path.resolve()), as_attachment=True)


@app.route("/download-combined/<job_id>")
def download_combined(job_id):
    """Download all converted files as a single combined markdown."""
    combined_path = OUTPUT_DIR / job_id / "all_combined.md"
    if not combined_path.exists():
        return jsonify({"error": "Combined file not ready yet"}), 404
    return send_file(str(combined_path.resolve()), as_attachment=True, download_name="all_combined.md")


@app.route("/download-merged/<job_id>")
def download_merged(job_id):
    """Download merged markdown file."""
    merged_path = OUTPUT_DIR / job_id / "merged_markdown.md"
    if not merged_path.exists():
        return jsonify({"error": "Merged file not found"}), 404
    return send_file(str(merged_path.resolve()), as_attachment=True, download_name="merged_markdown.md")


@app.route("/preview/<job_id>/<filename>")
def preview(job_id, filename):
    """Get markdown content for preview."""
    file_path = OUTPUT_DIR / job_id / filename
    if not file_path.exists():
        return jsonify({"error": "File not found"}), 404
    content = file_path.read_text(encoding="utf-8")
    # Truncate for preview
    if len(content) > 50000:
        content = content[:50000] + "\n\n... [truncated for preview] ..."
    return jsonify({"content": content, "filename": filename})


# ─── HTML Template ───────────────────────────────────────────────────────────

HTML_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PDF y Word → Markdown · Batch Converter</title>
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600;700&family=Outfit:wght@300;400;600;700;900&display=swap" rel="stylesheet">
<style>
:root {
    --bg: #0a0a0f;
    --surface: #12121a;
    --surface2: #1a1a26;
    --border: #2a2a3a;
    --text: #e8e8f0;
    --text-dim: #8888a0;
    --accent: #ff6b35;
    --accent2: #ff8f5e;
    --green: #34d399;
    --red: #f87171;
    --blue: #60a5fa;
    --yellow: #fbbf24;
    --radius: 12px;
}
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
    font-family: 'Outfit', sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
    overflow-x: hidden;
}

/* ── Animated background ── */
body::before {
    content: '';
    position: fixed;
    top: -50%; left: -50%;
    width: 200%; height: 200%;
    background: radial-gradient(circle at 30% 20%, rgba(255,107,53,0.04) 0%, transparent 50%),
                radial-gradient(circle at 70% 80%, rgba(96,165,250,0.03) 0%, transparent 50%);
    animation: drift 20s ease-in-out infinite;
    z-index: -1;
}
@keyframes drift {
    0%, 100% { transform: translate(0, 0); }
    50% { transform: translate(-3%, 2%); }
}

/* ── Header ── */
.header {
    padding: 2.5rem 2rem 1.5rem;
    text-align: center;
    position: relative;
}
.header h1 {
    font-size: 2.8rem;
    font-weight: 900;
    letter-spacing: -0.03em;
    background: linear-gradient(135deg, var(--accent), var(--accent2), var(--yellow));
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}
.header p {
    color: var(--text-dim);
    margin-top: 0.5rem;
    font-size: 1rem;
    font-weight: 300;
}
.badge {
    display: inline-block;
    margin-top: 0.75rem;
    padding: 0.3rem 0.8rem;
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: 20px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.72rem;
    color: var(--text-dim);
}

/* ── Container ── */
.container {
    max-width: 960px;
    margin: 0 auto;
    padding: 0 1.5rem 3rem;
}

/* ── Tabs Navigation ── */
.tabs-nav {
    display: flex;
    gap: 0.5rem;
    margin-bottom: 1.5rem;
    border-bottom: 1px solid var(--border);
    padding-bottom: 0;
}
.tab-btn {
    padding: 0.75rem 1.5rem;
    background: none;
    border: none;
    border-bottom: 2px solid transparent;
    color: var(--text-dim);
    cursor: pointer;
    font-family: 'Outfit', sans-serif;
    font-weight: 600;
    font-size: 0.95rem;
    transition: all 0.3s ease;
}
.tab-btn:hover {
    color: var(--text);
}
.tab-btn.active {
    color: var(--accent);
    border-bottom-color: var(--accent);
}

/* ── Tab Content ── */
.tab-content {
    display: none;
    animation: fadeIn 0.3s ease;
}
.tab-content.active {
    display: block;
}

/* ── Drop zone ── */
.dropzone {
    border: 2px dashed var(--border);
    border-radius: var(--radius);
    padding: 3.5rem 2rem;
    text-align: center;
    cursor: pointer;
    transition: all 0.3s ease;
    background: var(--surface);
    position: relative;
    overflow: hidden;
}
.dropzone::before {
    content: '';
    position: absolute;
    inset: 0;
    background: radial-gradient(circle at center, rgba(255,107,53,0.05), transparent 70%);
    opacity: 0;
    transition: opacity 0.3s;
}
.dropzone:hover, .dropzone.dragover {
    border-color: var(--accent);
    background: var(--surface2);
}
.dropzone:hover::before, .dropzone.dragover::before { opacity: 1; }
.dropzone-icon {
    font-size: 3rem;
    margin-bottom: 1rem;
    display: block;
}
.dropzone h3 {
    font-weight: 600;
    font-size: 1.2rem;
    margin-bottom: 0.5rem;
}
.dropzone p { color: var(--text-dim); font-size: 0.9rem; }

/* ── File list ── */
.file-queue {
    margin-top: 1.5rem;
    display: flex;
    flex-direction: column;
    gap: 0.5rem;
}
.file-item {
    display: grid;
    grid-template-columns: auto 1fr auto auto;
    align-items: center;
    gap: 0.75rem;
    padding: 0.85rem 1rem;
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    font-size: 0.88rem;
    transition: all 0.3s;
    animation: slideIn 0.3s ease;
}
@keyframes slideIn {
    from { opacity: 0; transform: translateY(-8px); }
    to { opacity: 1; transform: translateY(0); }
}
.file-item .icon { font-size: 1.3rem; }
.file-item .name {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.82rem;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.file-item .size {
    color: var(--text-dim);
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.75rem;
    white-space: nowrap;
}
.file-item .status-badge {
    padding: 0.2rem 0.6rem;
    border-radius: 20px;
    font-size: 0.72rem;
    font-weight: 600;
    white-space: nowrap;
}
.status-queued { background: var(--surface2); color: var(--text-dim); }
.status-converting {
    background: rgba(251,191,36,0.15);
    color: var(--yellow);
    animation: pulse 1.5s infinite;
}
@keyframes pulse {
    0%, 100% { opacity: 1; }
    50% { opacity: 0.6; }
}
.status-done { background: rgba(52,211,153,0.15); color: var(--green); }
.status-error { background: rgba(248,113,113,0.15); color: var(--red); }

.file-item .remove-btn {
    background: none;
    border: none;
    color: var(--text-dim);
    cursor: pointer;
    font-size: 1.1rem;
    padding: 0.2rem;
    border-radius: 4px;
    transition: all 0.2s;
}
.file-item .remove-btn:hover { color: var(--red); background: rgba(248,113,113,0.1); }

/* ── Actions ── */
.actions {
    margin-top: 1.5rem;
    display: flex;
    gap: 1rem;
    align-items: center;
    flex-wrap: wrap;
}
.btn {
    padding: 0.75rem 1.8rem;
    border: none;
    border-radius: 8px;
    font-family: 'Outfit', sans-serif;
    font-weight: 600;
    font-size: 0.95rem;
    cursor: pointer;
    transition: all 0.3s ease;
    display: inline-flex;
    align-items: center;
    gap: 0.5rem;
}
.btn-primary {
    background: linear-gradient(135deg, var(--accent), var(--accent2));
    color: #fff;
    box-shadow: 0 4px 20px rgba(255,107,53,0.25);
}
.btn-primary:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 30px rgba(255,107,53,0.35);
}
.btn-primary:disabled {
    opacity: 0.4;
    cursor: not-allowed;
    transform: none;
    box-shadow: none;
}
.btn-secondary {
    background: var(--surface2);
    color: var(--text);
    border: 1px solid var(--border);
}
.btn-secondary:hover { border-color: var(--text-dim); }
.btn-download {
    background: rgba(52,211,153,0.1);
    color: var(--green);
    border: 1px solid rgba(52,211,153,0.3);
}
.btn-download:hover {
    background: rgba(52,211,153,0.2);
    border-color: var(--green);
}

/* ── Progress ── */
.progress-container {
    margin-top: 1.5rem;
    display: none;
}
.progress-container.active { display: block; }
.progress-bar-bg {
    height: 6px;
    background: var(--surface2);
    border-radius: 3px;
    overflow: hidden;
}
.progress-bar {
    height: 100%;
    background: linear-gradient(90deg, var(--accent), var(--green));
    border-radius: 3px;
    transition: width 0.5s ease;
    width: 0%;
}
.progress-info {
    display: flex;
    justify-content: space-between;
    margin-top: 0.5rem;
    font-size: 0.82rem;
    color: var(--text-dim);
    font-family: 'JetBrains Mono', monospace;
}

/* ── Results ── */
.results-actions {
    margin-top: 1.5rem;
    display: none;
    gap: 0.75rem;
    flex-wrap: wrap;
}
.results-actions.active { display: flex; }

/* ── Preview modal ── */
.modal-overlay {
    display: none;
    position: fixed;
    inset: 0;
    background: rgba(0,0,0,0.7);
    z-index: 1000;
    animation: fadeIn 0.2s;
}
.modal-overlay.active { display: flex; align-items: center; justify-content: center; }
@keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
.modal {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    width: 90%;
    max-width: 800px;
    max-height: 80vh;
    display: flex;
    flex-direction: column;
    animation: scaleIn 0.2s ease;
}
@keyframes scaleIn {
    from { transform: scale(0.95); opacity: 0; }
    to { transform: scale(1); opacity: 1; }
}
.modal-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 1rem 1.5rem;
    border-bottom: 1px solid var(--border);
}
.modal-header h3 {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.9rem;
    font-weight: 600;
}
.modal-close {
    background: none;
    border: none;
    color: var(--text-dim);
    font-size: 1.5rem;
    cursor: pointer;
}
.modal-body {
    padding: 1.5rem;
    overflow-y: auto;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.8rem;
    line-height: 1.7;
    white-space: pre-wrap;
    color: var(--text-dim);
}

/* ── Stats ── */
.stats-bar {
    display: none;
    grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
    gap: 0.75rem;
    margin-top: 1.5rem;
}
.stats-bar.active { display: grid; }
.stat-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 1rem;
    text-align: center;
}
.stat-card .value {
    font-size: 1.5rem;
    font-weight: 700;
    font-family: 'JetBrains Mono', monospace;
}
.stat-card .label {
    font-size: 0.75rem;
    color: var(--text-dim);
    margin-top: 0.25rem;
}

/* ── Options panel ── */
.options-panel {
    margin-top: 1rem;
    padding: 1rem 1.25rem;
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    display: none;
}
.options-panel.active { display: block; }
.options-panel h4 {
    font-size: 0.85rem;
    margin-bottom: 0.75rem;
    color: var(--text-dim);
}
.option-row {
    display: flex;
    align-items: center;
    gap: 0.75rem;
    margin-bottom: 0.5rem;
}
.option-row label {
    font-size: 0.85rem;
    cursor: pointer;
    display: flex;
    align-items: center;
    gap: 0.4rem;
}
.option-row input[type="checkbox"] {
    accent-color: var(--accent);
    width: 16px;
    height: 16px;
}

/* ── Hidden input ── */
#fileInput { display: none; }

/* ── Responsive ── */
@media (max-width: 600px) {
    .header h1 { font-size: 2rem; }
    .dropzone { padding: 2rem 1rem; }
    .file-item { grid-template-columns: auto 1fr auto; }
    .file-item .size { display: none; }
}
</style>
</head>
<body>

<div class="header">
    <h1>PDF y Word → Markdown</h1>
    <p>Conversión en lote · Local · Sin límites</p>
    <span class="badge">PDF · Word .docx · tablas · imágenes · multi-columna</span>
</div>

<div class="container">

    <!-- Tabs Navigation -->
    <div class="tabs-nav" id="tabsNav">
        <button class="tab-btn active" data-tab="pdf-panel" onclick="switchTab('pdf-panel')">
            📄 Convertir archivos
        </button>
        <button class="tab-btn" data-tab="merge-panel" onclick="switchTab('merge-panel')">
            📝 Unir MDs
        </button>
    </div>

    <!-- PDF Conversion Panel -->
    <div class="tab-content active" id="pdf-panel">

        <!-- Drop Zone -->
        <div class="dropzone" id="dropzone" onclick="document.getElementById('fileInput').click()">
            <span class="dropzone-icon">📄</span>
            <h3>Arrastra tus archivos PDF o Word aquí</h3>
            <p>o haz clic para seleccionar · acepta .pdf, .docx y .docm</p>
        </div>
        <input type="file" id="fileInput" accept=".pdf,.docx,.docm" multiple>

        <!-- Options toggle -->
        <div style="margin-top: 0.75rem; text-align: right;">
            <button class="btn btn-secondary" onclick="toggleOptions()" style="padding: 0.4rem 1rem; font-size: 0.8rem;">
                ⚙ Opciones
            </button>
        </div>
        <div class="options-panel" id="optionsPanel">
            <h4>Opciones de conversión</h4>
            <div class="option-row">
                <label><input type="checkbox" id="optImages" checked> Extraer imágenes (PNG)</label>
            </div>
            <div class="option-row">
                <label><input type="checkbox" id="optHeaders" checked> Excluir encabezados/pies repetitivos</label>
            </div>
            <div class="option-row" style="margin-top: 0.8rem; display: block;">
                <h4 style="margin-bottom: 0.4rem;">Reconocimiento de Texto (OCR)</h4>
                <select id="ocrMode" style="width: 100%; padding: 0.5rem; background: var(--surface2); border: 1px solid var(--border); color: var(--text); border-radius: 6px; font-family: 'Outfit', sans-serif; cursor: pointer;">
                    <option value="auto">Auto-detectar (Recomendado)</option>
                    <option value="force">Forzar OCR (Especial para escaneos)</option>
                    <option value="none">Desactivado (Más rápido, solo digital)</option>
                </select>
                <div style="display:flex;gap:0.6rem;margin-top:0.6rem;align-items:center;">
                    <label style="flex:1; font-size:0.85rem; color:var(--text-dim);">DPI para OCR:</label>
                    <input id="ocrDpi" type="number" min="72" max="600" value="300" style="width:110px;padding:0.4rem;border-radius:6px;background:var(--surface2);border:1px solid var(--border);color:var(--text);" />
                    <label style="margin-left:0.6rem; font-size:0.85rem;"><input id="ocrPreprocess" type="checkbox" style="margin-right:0.35rem;"> Mejorar imágenes (preprocesar)</label>
                </div>
                <p style="font-size: 0.72rem; color: var(--text-dim); margin-top: 0.35rem; line-height: 1.4;">
                    <strong>Auto</strong> solo activa el motor de OCR si el documento parece estar escaneado. Aumentar DPI mejora la precisión del OCR, pero aumenta el tiempo.
                </p>
            </div>
        </div>

        <!-- File Queue -->
        <div class="file-queue" id="fileQueue"></div>

        <!-- Actions -->
        <div class="actions" id="actionsBar" style="display: none;">
            <button class="btn btn-primary" id="convertBtn" onclick="startConversion()">
                ⚡ Convertir todo
            </button>
            <button class="btn btn-secondary" onclick="clearAll()">
                ✕ Limpiar
            </button>
            <span id="fileCount" style="color: var(--text-dim); font-size: 0.85rem; font-family: 'JetBrains Mono', monospace;"></span>
        </div>

        <!-- Progress -->
        <div class="progress-container" id="progressContainer">
            <div class="progress-bar-bg">
                <div class="progress-bar" id="progressBar"></div>
            </div>
            <div class="progress-info">
                <span id="progressText">0 / 0 archivos</span>
                <span id="progressPercent">0%</span>
            </div>
        </div>

        <!-- Stats -->
        <div class="stats-bar" id="statsBar">
            <div class="stat-card">
                <div class="value" id="statTotal">0</div>
                <div class="label">Total archivos</div>
            </div>
            <div class="stat-card">
                <div class="value" id="statDone" style="color: var(--green);">0</div>
                <div class="label">Convertidos</div>
            </div>
            <div class="stat-card">
                <div class="value" id="statErrors" style="color: var(--red);">0</div>
                <div class="label">Errores</div>
            </div>
            <div class="stat-card">
                <div class="value" id="statTime">0s</div>
                <div class="label">Tiempo total</div>
            </div>
        </div>

        <!-- Results actions -->
        <div class="results-actions" id="resultsActions">
            <div style="width: 100%; margin-bottom: 1rem; border-top: 1px solid var(--border); padding-top: 1.5rem;">
                <h3 style="font-size: 1.1rem; margin-bottom: 1rem; color: var(--accent);">✨ ¡Conversión completada! Elige cómo guardar:</h3>
            </div>
            <button class="btn btn-download" id="downloadCombinedBtn" onclick="downloadCombined()" style="flex: 1; padding: 1rem; border-radius: 12px; font-size: 1.05rem;">
                📝 <strong>Descargar MD unido</strong><br>
                <span style="font-size: 0.75rem; opacity: 0.8;">Todos los archivos en uno solo</span>
            </button>
            <button class="btn btn-download" id="downloadAllBtn" onclick="downloadAll()" style="flex: 1; padding: 1rem; border-radius: 12px; font-size: 1.05rem;">
                📦 <strong>Descargar todo (.zip)</strong><br>
                <span style="font-size: 0.75rem; opacity: 0.8;">Archivos individuales e imágenes</span>
            </button>
        </div>

    </div>

    <!-- Merge MD Panel -->
    <div class="tab-content" id="merge-panel">

        <!-- Drop Zone for MD files -->
        <div class="dropzone" id="dropzoneMd" onclick="document.getElementById('fileInputMd').click()">
            <span class="dropzone-icon">📝</span>
            <h3>Arrastra tus archivos Markdown aquí</h3>
            <p>o haz clic para seleccionar · acepta múltiples archivos .md</p>
        </div>
        <input type="file" id="fileInputMd" accept=".md" multiple>

        <!-- File Queue for MD -->
        <div class="file-queue" id="fileQueueMd"></div>

        <!-- Actions for MD -->
        <div class="actions" id="actionsBarMd" style="display: none;">
            <button class="btn btn-primary" id="mergeBtn" onclick="startMerge()">
                🔗 Unir MDs
            </button>
            <button class="btn btn-secondary" onclick="clearAllMd()">
                ✕ Limpiar
            </button>
            <span id="fileCountMd" style="color: var(--text-dim); font-size: 0.85rem; font-family: 'JetBrains Mono', monospace;"></span>
        </div>

        <!-- Merge Progress -->
        <div class="progress-container" id="mergeProgressContainer">
            <div class="progress-bar-bg">
                <div class="progress-bar" id="mergeProgressBar"></div>
            </div>
            <div class="progress-info">
                <span id="mergeProgressText">Procesando...</span>
            </div>
        </div>

        <!-- Merge Results -->
        <div class="results-actions" id="mergeMergeResults">
            <div style="width: 100%; margin-bottom: 1rem; border-top: 1px solid var(--border); padding-top: 1.5rem;">
                <h3 style="font-size: 1.1rem; margin-bottom: 1rem; color: var(--accent);">✨ ¡Unión completada!</h3>
            </div>
            <button class="btn btn-download" id="downloadMergedBtn" onclick="downloadMerged()" style="flex: 1; padding: 1rem; border-radius: 12px; font-size: 1.05rem;">
                📥 <strong>Descargar archivo único</strong><br>
                <span style="font-size: 0.75rem; opacity: 0.8;">Todo el contenido unido en un MD</span>
            </button>
        </div>

    </div>

</div>

<!-- Preview Modal -->
<div class="modal-overlay" id="previewModal">
    <div class="modal">
        <div class="modal-header">
            <h3 id="previewTitle">preview.md</h3>
            <button class="modal-close" onclick="closePreview()">×</button>
        </div>
        <div class="modal-body" id="previewBody"></div>
    </div>
</div>

<script>
// ── State ──
let pendingFiles = [];
let pendingMdFiles = [];
let currentJobId = null;
let pollInterval = null;

const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const fileQueue = document.getElementById('fileQueue');

const dropzoneMd = document.getElementById('dropzoneMd');
const fileInputMd = document.getElementById('fileInputMd');
const fileQueueMd = document.getElementById('fileQueueMd');

// ── Tab Switching ──
function switchTab(tabName) {
    // Hide all tabs
    document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));

    // Show selected tab
    document.getElementById(tabName).classList.add('active');
    document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');
}

// ── Drag & Drop PDF ──
['dragenter', 'dragover'].forEach(e => {
    dropzone.addEventListener(e, ev => { ev.preventDefault(); dropzone.classList.add('dragover'); });
});
['dragleave', 'drop'].forEach(e => {
    dropzone.addEventListener(e, ev => { ev.preventDefault(); dropzone.classList.remove('dragover'); });
});
dropzone.addEventListener('drop', ev => {
    const files = [...ev.dataTransfer.files].filter(isConvertibleFile);
    addFiles(files);
});
fileInput.addEventListener('change', () => {
    addFiles([...fileInput.files]);
    fileInput.value = '';
});

// ── Drag & Drop MD ──
['dragenter', 'dragover'].forEach(e => {
    dropzoneMd.addEventListener(e, ev => { ev.preventDefault(); dropzoneMd.classList.add('dragover'); });
});
['dragleave', 'drop'].forEach(e => {
    dropzoneMd.addEventListener(e, ev => { ev.preventDefault(); dropzoneMd.classList.remove('dragover'); });
});
dropzoneMd.addEventListener('drop', ev => {
    const files = [...ev.dataTransfer.files].filter(f => f.name.toLowerCase().endsWith('.md'));
    addMdFiles(files);
});
fileInputMd.addEventListener('change', () => {
    addMdFiles([...fileInputMd.files]);
    fileInputMd.value = '';
});

// ── File Management ──
function addFiles(files) {
    files.forEach(f => {
        if (!pendingFiles.find(p => p.name === f.name && p.size === f.size)) {
            pendingFiles.push(f);
        }
    });
    renderQueue();
}

function removeFile(index) {
    pendingFiles.splice(index, 1);
    renderQueue();
}

function clearAll() {
    pendingFiles = [];
    currentJobId = null;
    if (pollInterval) clearInterval(pollInterval);
    renderQueue();
    document.getElementById('progressContainer').classList.remove('active');
    document.getElementById('statsBar').classList.remove('active');
    document.getElementById('resultsActions').classList.remove('active');
}

// ── MD File Management ──
function addMdFiles(files) {
    files.forEach(f => {
        if (!pendingMdFiles.find(p => p.name === f.name && p.size === f.size)) {
            pendingMdFiles.push(f);
        }
    });
    renderMdQueue();
}

function removeMdFile(index) {
    pendingMdFiles.splice(index, 1);
    renderMdQueue();
}

function clearAllMd() {
    pendingMdFiles = [];
    renderMdQueue();
    document.getElementById('mergeProgressContainer').classList.remove('active');
    document.getElementById('mergeMergeResults').classList.remove('active');
}

function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function isConvertibleFile(file) {
    return /\.(pdf|docx|docm)$/i.test(file.name);
}

function fileIcon(filename) {
    return /\.(docx|docm)$/i.test(filename) ? '\u{1F4DD}' : '\u{1F4C4}';
}

function renderQueue() {
    fileQueue.innerHTML = '';
    pendingFiles.forEach((f, i) => {
        const div = document.createElement('div');
        div.className = 'file-item';
        div.innerHTML = `
            <span class="icon">📄</span>
            <span class="name" title="${f.name}">${f.name}</span>
            <span class="size">${formatSize(f.size)}</span>
            <button class="remove-btn" onclick="removeFile(${i})" title="Quitar">×</button>
        `;
        fileQueue.appendChild(div);
    });

    const bar = document.getElementById('actionsBar');
    const count = document.getElementById('fileCount');
    if (pendingFiles.length > 0) {
        bar.style.display = 'flex';
        count.textContent = `${pendingFiles.length} archivo${pendingFiles.length > 1 ? 's' : ''} listo${pendingFiles.length > 1 ? 's' : ''}`;
    } else {
        bar.style.display = 'none';
    }
}

function renderMdQueue() {
    fileQueueMd.innerHTML = '';
    pendingMdFiles.forEach((f, i) => {
        const div = document.createElement('div');
        div.className = 'file-item';
        div.innerHTML = `
            <span class="icon">📝</span>
            <span class="name" title="${f.name}">${f.name}</span>
            <span class="size">${formatSize(f.size)}</span>
            <button class="remove-btn" onclick="removeMdFile(${i})" title="Quitar">×</button>
        `;
        fileQueueMd.appendChild(div);
    });

    const bar = document.getElementById('actionsBarMd');
    const count = document.getElementById('fileCountMd');
    if (pendingMdFiles.length > 0) {
        bar.style.display = 'flex';
        count.textContent = `${pendingMdFiles.length} archivo${pendingMdFiles.length > 1 ? 's' : ''} listo${pendingMdFiles.length > 1 ? 's' : ''}`;
    } else {
        bar.style.display = 'none';
    }
}

// ── Conversion ──
async function startConversion() {
    if (pendingFiles.length === 0) return;

    const btn = document.getElementById('convertBtn');
    btn.disabled = true;
    btn.innerHTML = '⏳ Subiendo...';

    const formData = new FormData();
    pendingFiles.forEach(f => formData.append('files', f));

    // Add options
    formData.append('optImages', document.getElementById('optImages').checked);
    formData.append('optHeaders', document.getElementById('optHeaders').checked);
    formData.append('ocrMode', document.getElementById('ocrMode').value);
    formData.append('ocrDpi', document.getElementById('ocrDpi').value);
    formData.append('ocrPreprocess', document.getElementById('ocrPreprocess').checked);

    try {
        const res = await fetch('/upload', { method: 'POST', body: formData });
        const data = await res.json();
        if (data.error) {
            alert('Error: ' + data.error);
            btn.disabled = false;
            btn.innerHTML = '⚡ Convertir todo';
            return;
        }

        currentJobId = data.job_id;
        btn.innerHTML = '⏳ Convirtiendo...';
        document.getElementById('progressContainer').classList.add('active');

        // Start polling
        pollInterval = setInterval(pollStatus, 800);

    } catch (err) {
        alert('Error de conexión: ' + err.message);
        btn.disabled = false;
        btn.innerHTML = '⚡ Convertir todo';
    }
}

// ── Merge MD ──
async function startMerge() {
    if (pendingMdFiles.length === 0) return;

    const btn = document.getElementById('mergeBtn');
    btn.disabled = true;
    btn.innerHTML = '⏳ Uniendo...';

    const formData = new FormData();
    pendingMdFiles.forEach(f => formData.append('mds', f));

    try {
        const res = await fetch('/merge-md', { method: 'POST', body: formData });
        const data = await res.json();
        if (data.error) {
            alert('Error: ' + data.error);
            btn.disabled = false;
            btn.innerHTML = '🔗 Unir MDs';
            return;
        }

        document.getElementById('mergeProgressContainer').classList.add('active');
        document.getElementById('mergeProgressText').textContent = `✓ ${data.total} archivos unidos exitosamente`;
        btn.innerHTML = '✓ Completado';
        document.getElementById('mergeMergeResults').classList.add('active');

        // Store the merge job ID for download
        window.mergeJobId = data.job_id;

    } catch (err) {
        alert('Error de conexión: ' + err.message);
        btn.disabled = false;
        btn.innerHTML = '🔗 Unir MDs';
    }
}

async function pollStatus() {
    if (!currentJobId) return;
    try {
        const res = await fetch(`/status/${currentJobId}`);
        const job = await res.json();
        updateUI(job);
        if (job.status === 'done') {
            clearInterval(pollInterval);
            document.getElementById('convertBtn').innerHTML = '✓ Completado';
            document.getElementById('resultsActions').classList.add('active');
        }
    } catch (e) { /* ignore transient errors */ }
}

function updateUI(job) {
    const pct = job.total > 0 ? Math.round((job.completed / job.total) * 100) : 0;
    document.getElementById('progressBar').style.width = pct + '%';
    document.getElementById('progressText').textContent = `${job.completed} / ${job.total} archivos`;
    document.getElementById('progressPercent').textContent = pct + '%';

    // Update file items
    fileQueue.innerHTML = '';
    job.files.forEach((f, i) => {
        const div = document.createElement('div');
        div.className = 'file-item';

        let statusClass = 'status-' + f.status;
        let statusLabel = {
            'queued': 'En cola',
            'converting': 'Convirtiendo…',
            'done': '✓ Listo',
            'error': '✗ Error'
        }[f.status] || f.status;

        let actions = '';
        if (f.status === 'done' && f.output) {
            actions = `
                <a href="/download/${job.id}/${f.output}" class="btn btn-secondary" style="padding:0.25rem 0.6rem;font-size:0.72rem;text-decoration:none;">↓ .md</a>
                <button class="btn btn-secondary" style="padding:0.25rem 0.6rem;font-size:0.72rem;" onclick="showPreview('${job.id}','${f.output}')">👁</button>
            `;
        }
        if (f.status === 'error') {
            actions = `<span style="color:var(--red);font-size:0.72rem;" title="${f.error || ''}">ver error</span>`;
        }

        let timeInfo = f.time ? `${f.time}s` : '';

        div.innerHTML = `
            <span class="icon">📄</span>
            <span class="name" title="${f.original_name}">${f.original_name}</span>
            <span class="status-badge ${statusClass}">${statusLabel}${timeInfo ? ' · ' + timeInfo : ''}</span>
            <span style="display:flex;gap:0.3rem;align-items:center;">${actions}</span>
        `;
        fileQueue.appendChild(div);
    });

    // Stats
    if (job.completed > 0) {
        document.getElementById('statsBar').classList.add('active');
        document.getElementById('statTotal').textContent = job.total;
        const done = job.files.filter(f => f.status === 'done').length;
        const errors = job.files.filter(f => f.status === 'error').length;
        const totalTime = job.files.reduce((s, f) => s + (f.time || 0), 0);
        document.getElementById('statDone').textContent = done;
        document.getElementById('statErrors').textContent = errors;
        document.getElementById('statTime').textContent = totalTime.toFixed(1) + 's';
    }
}

// ── Download All ──
function downloadAll() {
    if (currentJobId) {
        window.location.href = `/download-all/${currentJobId}`;
    }
}

function downloadCombined() {
    if (currentJobId) {
        window.location.href = `/download-combined/${currentJobId}`;
    }
}

function downloadMerged() {
    if (window.mergeJobId) {
        window.location.href = `/download-merged/${window.mergeJobId}`;
    }
}

// ── Preview ──
async function showPreview(jobId, filename) {
    const res = await fetch(`/preview/${jobId}/${filename}`);
    const data = await res.json();
    document.getElementById('previewTitle').textContent = data.filename;
    document.getElementById('previewBody').textContent = data.content;
    document.getElementById('previewModal').classList.add('active');
}
function closePreview() {
    document.getElementById('previewModal').classList.remove('active');
}
document.getElementById('previewModal').addEventListener('click', e => {
    if (e.target === document.getElementById('previewModal')) closePreview();
});

// ── Options toggle ──
function toggleOptions() {
    document.getElementById('optionsPanel').classList.toggle('active');
}
</script>
</body>
</html>
"""


# ─── Entry point ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 5000
    print(f"""
+----------------------------------------------------+
|   PDF -> Markdown  ·  Batch Converter             |
|   http://localhost:{port}                          |
|   Ctrl+C para detener                            |
+----------------------------------------------------+
    """)
    app.run(host="0.0.0.0", port=port, debug=False)
