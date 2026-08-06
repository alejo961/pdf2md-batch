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
import re
import time
import zipfile
import shutil
import threading
from pathlib import Path
from datetime import datetime
from uuid import uuid4
from concurrent.futures import ThreadPoolExecutor

from flask import (
    Flask, render_template_string, request, jsonify,
    send_file, send_from_directory
)

import pymupdf4llm
import fitz

from conversion_formats import (
    markdown_to_docx, markdown_to_pdf,
)

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


def configure_tesseract() -> Path | None:
    """Locate a local Tesseract installation so PyMuPDF OCR can use it."""
    configured = os.environ.get("TESSDATA_PREFIX")
    candidates = [
        Path(configured) if configured else None,
        Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Tesseract-OCR" / "tessdata",
        Path(os.environ.get("PROGRAMFILES", "")) / "Tesseract-OCR" / "tessdata",
        Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Tesseract-OCR" / "tessdata",
    ]
    for candidate in candidates:
        if candidate and candidate.is_dir() and (candidate / "eng.traineddata").is_file():
            os.environ["TESSDATA_PREFIX"] = str(candidate)
            return candidate
    return None


TESSDATA_DIR = configure_tesseract()

# Track conversion jobs
jobs = {}
jobs_lock = threading.Lock()
ocr_lock = threading.Lock()
OCR_MAX_ATTEMPTS = 3
OCR_RETRY_DELAY_SECONDS = 0.5


def persist_job_snapshot_locked(job_id: str) -> None:
    """Atomically persist a job while jobs_lock is held."""
    job = jobs.get(job_id)
    if job is None:
        return
    job_dir = OUTPUT_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    snapshot_path = job_dir / "job.json"
    temporary_path = job_dir / "job.json.tmp"
    temporary_path.write_text(
        json.dumps(job, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    os.replace(temporary_path, snapshot_path)


def load_job_snapshots() -> list[str]:
    """Load persisted jobs and return interrupted job IDs that can resume."""
    resumable = []
    for snapshot_path in OUTPUT_DIR.glob("*/job.json"):
        try:
            job = json.loads(snapshot_path.read_text(encoding="utf-8"))
            job_id = str(job.get("id") or snapshot_path.parent.name)
            if not JOB_ID_PATTERN.fullmatch(job_id):
                continue
            if job.get("status") == "running":
                for item in job.get("files", []):
                    if item.get("status") in {"queued", "converting"}:
                        item["status"] = "queued"
                job["completed"] = sum(
                    item.get("status") in {"done", "error"}
                    for item in job.get("files", [])
                )
                if (UPLOAD_DIR / job_id).is_dir():
                    resumable.append(job_id)
            jobs[job_id] = job
        except (OSError, ValueError, TypeError):
            continue
    return resumable
CONVERSION_MAX_WORKERS = min(os.cpu_count() or 4, 4)
conversion_executor = ThreadPoolExecutor(max_workers=CONVERSION_MAX_WORKERS)
JOB_ID_PATTERN = re.compile(r"^[A-Za-z0-9_-]+$")
OMITTED_IMAGE_PATTERN = re.compile(
    r"(?mi)^\s*(?:\*{0,2})?==>\s*picture\s*\[[^\]]*\]\s*"
    r"intentionally omitted\s*<==(?:\*{0,2})?\s*$"
)


def sanitize_upload_name(filename: str) -> str:
    """Return a portable filename that cannot escape its job directory."""
    name = Path(filename.replace("\\", "/")).name
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")
    return name or "document"

def unique_upload_name(filename: str, used_names: set[str]) -> str:
    """Avoid concurrent writes when uploaded filenames collide."""
    candidate = sanitize_upload_name(filename)
    stem, suffix = Path(candidate).stem, Path(candidate).suffix
    counter = 2
    while candidate.casefold() in used_names:
        candidate = f"{stem}_{counter}{suffix}"
        counter += 1
    used_names.add(candidate.casefold())
    return candidate


def resolve_job_directory(root: Path, job_id: str) -> Path | None:
    """Resolve a job directory only when it remains below the configured root."""
    if not JOB_ID_PATTERN.fullmatch(job_id or ""):
        return None
    base = root.resolve()
    candidate = (base / job_id).resolve()
    try:
        candidate.relative_to(base)
    except ValueError:
        return None
    return candidate


def resolve_job_file(root: Path, job_id: str, filename: str) -> Path | None:
    """Resolve one direct child file of a validated job directory."""
    job_dir = resolve_job_directory(root, job_id)
    if job_dir is None or sanitize_upload_name(filename) != filename:
        return None
    candidate = (job_dir / filename).resolve()
    try:
        candidate.relative_to(job_dir)
    except ValueError:
        return None
    return candidate


def clean_page_markdown(markdown_text: str) -> str:
    """Remove extractor placeholders and empty Markdown headings."""
    text = OMITTED_IMAGE_PATTERN.sub("", markdown_text or "")
    text = re.sub(r"(?m)^#{1,6}\s*$", "", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def markdown_semantic_word_count(markdown_text: str) -> int:
    """Count meaningful words after removing Markdown-only structure."""
    text = clean_page_markdown(markdown_text)
    text = re.sub(r"!\[[^\]]*\]\([^)]*\)|!\[\[[^\]]+\]\]", "", text)
    text = re.sub(r"(?m)^## Página \d+\s*$|^[-_=]{3,}\s*$", "", text)
    text = re.sub(r"[#*_`>|]", " ", text)
    return len(re.findall(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9]{3,}", text))


# ─── Utility Logic ───────────────────────────────────────────────────────────

def page_has_usable_text(page, minimum_characters: int = 40) -> bool:
    """Return whether a PDF page has enough native text to skip OCR."""
    text = page.get_text("text") or ""
    compact_text = re.sub(r"\s+", "", text)
    return len(compact_text) >= minimum_characters
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


def is_image_only_markdown(markdown_text: str) -> bool:
    """Detect output that contains image references but no useful text."""
    stripped = (markdown_text or "").strip()
    if not stripped:
        return True
    has_images = bool(re.search(r"!\[[^\]]*\]\([^)]*\)|!\[\[[^\]]+\]\]", stripped))
    return has_images and markdown_semantic_word_count(stripped) < 10

def ocr_page_to_markdown(page, page_number: int, dpi: int) -> str:
    """Run thread-safe OCR on one page, retrying temporary Leptonica failures."""
    last_error = None

    for attempt in range(1, OCR_MAX_ATTEMPTS + 1):
        try:
            # PyMuPDF invokes Leptonica in-process. Leptonica rejects concurrent
            # calls, so only the OCR section is serialized; digital extraction
            # and the rest of each PDF conversion remain parallel.
            with ocr_lock:
                try:
                    textpage = page.get_textpage_ocr(
                        language="spa+eng", dpi=dpi, full=True
                    )
                except Exception as spanish_error:
                    if "Leptonica from 2 threads" in str(spanish_error):
                        raise
                    textpage = page.get_textpage_ocr(
                        language="eng", dpi=dpi, full=True
                    )
                return page.get_text("text", textpage=textpage).strip()
        except Exception as exc:
            last_error = exc
            is_temporary = "Leptonica from 2 threads" in str(exc)
            if not is_temporary or attempt == OCR_MAX_ATTEMPTS:
                break
            time.sleep(OCR_RETRY_DELAY_SECONDS * attempt)

    raise RuntimeError(
        f"No se pudo hacer OCR en la página {page_number}. Tesseract está instalado, "
        "pero el motor OCR no pudo procesar esta página. Detalle: " + str(last_error)
    ) from last_error

def pdf_to_markdown_page_by_page(pdf_path: Path, output_path: Path, options: dict) -> str:
    """Convert every page independently, applying OCR only where required."""
    ocr_mode = options.get("ocr_mode", "auto")
    dpi = max(72, min(int(options.get("ocr_dpi") or 300), 600))
    write_images = bool(options.get("extract_images", False))
    page_blocks = []

    with fitz.open(str(pdf_path)) as doc:
        for page_index, page in enumerate(doc):
            page_number = page_index + 1
            native_text = (page.get_text("text") or "").strip()
            page_has_images = bool(page.get_images(full=True))
            should_ocr = ocr_mode == "force" or (
                ocr_mode == "auto"
                and not page_has_usable_text(page)
                and page_has_images
            )

            page_markdown = ""
            if not should_ocr:
                page_markdown = pymupdf4llm.to_markdown(
                    doc,
                    pages=[page_index],
                    write_images=write_images,
                    image_path=str(output_path.parent / "images"),
                    image_format="png",
                    dpi=min(dpi, 300),
                    use_ocr=False,
                    header=not options.get("exclude_headers", False),
                    footer=not options.get("exclude_headers", False),
                )
                page_markdown = clean_page_markdown(page_markdown)

                if is_image_only_markdown(page_markdown):
                    if ocr_mode == "none":
                        raise RuntimeError(
                            f"La página {page_number} solo produjo referencias a imágenes. "
                            "Activa OCR automático o forzado para extraer su texto."
                        )
                    should_ocr = True
                elif (
                    markdown_semantic_word_count(page_markdown) < 10
                    and markdown_semantic_word_count(native_text) >= 10
                ):
                    # Some PDFs expose valid native text while the layout engine
                    # returns only empty headings. Preserve that native text.
                    page_markdown = native_text
                elif (
                    markdown_semantic_word_count(page_markdown) == 0
                    and page_has_images
                    and ocr_mode != "none"
                ):
                    should_ocr = True

            if should_ocr:
                page_markdown = clean_page_markdown(
                    ocr_page_to_markdown(page, page_number, dpi)
                )
                if not page_markdown:
                    if page_has_images:
                        raise RuntimeError(
                            f"OCR no devolvió texto en la página {page_number}. "
                            "Verifica la calidad del escaneo."
                        )
                    page_markdown = "_Página sin texto reconocible._"

            if not page_markdown:
                page_markdown = "_Página sin texto reconocible._"
            page_blocks.append(f"## Página {page_number}\n\n{page_markdown}".strip())

    markdown_text = "\n\n---\n\n".join(page_blocks).strip()
    if markdown_semantic_word_count(markdown_text) == 0:
        raise RuntimeError(
            "La conversión no produjo texto útil; no se creó un Markdown vacío."
        )
    return markdown_text + "\n"

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
    """Convert a PDF page by page, with per-page automatic OCR."""
    options = options or {}
    start_time = time.time()

    try:
        md_text = pdf_to_markdown_page_by_page(pdf_path, output_path, options)
        output_name = write_markdown_output(md_text, output_path)
        elapsed = round(time.time() - start_time, 2)

        with jobs_lock:
            jobs[job_id]["files"][file_index]["status"] = "done"
            jobs[job_id]["files"][file_index]["output"] = output_name
            jobs[job_id]["files"][file_index]["size_md"] = len(md_text)
            jobs[job_id]["files"][file_index]["time"] = elapsed
            jobs[job_id]["completed"] += 1
            persist_job_snapshot_locked(job_id)

    except Exception as e:
        with jobs_lock:
            jobs[job_id]["files"][file_index]["status"] = "error"
            jobs[job_id]["files"][file_index]["error"] = str(e)
            jobs[job_id]["completed"] += 1
            persist_job_snapshot_locked(job_id)


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

        output_name = write_markdown_output(md_text, output_path)
        elapsed = round(time.time() - start_time, 2)

        with jobs_lock:
            jobs[job_id]["files"][file_index]["status"] = "done"
            jobs[job_id]["files"][file_index]["output"] = output_name
            jobs[job_id]["files"][file_index]["size_md"] = len(md_text)
            jobs[job_id]["files"][file_index]["time"] = elapsed
            jobs[job_id]["completed"] += 1
            persist_job_snapshot_locked(job_id)

    except Exception as e:
        with jobs_lock:
            jobs[job_id]["files"][file_index]["status"] = "error"
            jobs[job_id]["files"][file_index]["error"] = str(e)
            jobs[job_id]["completed"] += 1
            persist_job_snapshot_locked(job_id)


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
            persist_job_snapshot_locked(job_id)


def write_markdown_output(markdown_text: str, output_path: Path) -> str:
    """Write one portable standard Markdown output."""
    output_path.write_text(markdown_text, encoding="utf-8")
    return output_path.name


def combine_md_files(job_output_dir: Path, output_names: list[str]) -> Path:
    """Combine the successful standard Markdown outputs into one file."""
    combined_path = job_output_dir / "all_combined.md"
    markdown_files = [
        job_output_dir / name
        for name in output_names
        if name and (job_output_dir / name).is_file()
    ]
    with combined_path.open("w", encoding="utf-8") as output_file:
        for index, markdown_file in enumerate(markdown_files):
            if index:
                output_file.write("\n\n" + "=" * 80 + "\n\n")
            output_file.write(markdown_file.read_text(encoding="utf-8"))
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


def convert_queued_file_to_md(
    input_path: Path,
    output_path: Path,
    job_id: str,
    file_index: int,
    options: dict,
):
    """Mark a queued file active only when a global worker actually starts it."""
    with jobs_lock:
        job = jobs.get(job_id)
        if job is None:
            return
        job["files"][file_index]["status"] = "converting"
        persist_job_snapshot_locked(job_id)
    convert_file_to_md(input_path, output_path, job_id, file_index, options)


def run_batch_conversion(job_id: str):
    """Process a job through the shared bounded conversion queue."""
    job = jobs[job_id]
    job_upload_dir = UPLOAD_DIR / job_id
    job_output_dir = OUTPUT_DIR / job_id
    job_output_dir.mkdir(parents=True, exist_ok=True)
    (job_output_dir / "images").mkdir(exist_ok=True)
    options = job.get("options", {})

    futures = []
    for i, file_info in enumerate(job["files"]):
        if file_info.get("status") in {"done", "error"}:
            continue
        input_path = job_upload_dir / file_info["original_name"]
        output_path = job_output_dir / (Path(file_info["original_name"]).stem + ".md")
        futures.append(
            conversion_executor.submit(
                convert_queued_file_to_md,
                input_path,
                output_path,
                job_id,
                i,
                options,
            )
        )

    for future in futures:
        try:
            future.result()
        except Exception as exc:
            print(f"Error in future result: {exc}")

    with jobs_lock:
        job["status"] = "packaging"
        persist_job_snapshot_locked(job_id)

    output_names = [
        item.get("output") for item in job["files"]
        if item.get("status") == "done" and item.get("output")
    ]
    combine_md_files(job_output_dir, output_names)

    zip_path = job_output_dir / "all_markdown.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in [*output_names, "all_combined.md"]:
            path = job_output_dir / name
            if path.is_file():
                zf.write(path, path.name)
        images_dir = job_output_dir / "images"
        if images_dir.exists():
            for image_path in images_dir.iterdir():
                zf.write(image_path, f"images/{image_path.name}")

    with jobs_lock:
        job["status"] = "done"
        job["finished_at"] = datetime.now().isoformat()
        persist_job_snapshot_locked(job_id)

def export_markdown_job(md_files, asset_files, export_pdf: bool, export_word: bool) -> dict:
    """Convert uploaded Markdown files to polished PDF and/or Word outputs."""
    job_id = datetime.now().strftime("%Y%m%d_%H%M%S_export_") + uuid4().hex[:8]
    upload_dir = UPLOAD_DIR / job_id
    output_dir = OUTPUT_DIR / job_id
    upload_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    for asset in asset_files:
        if not asset.filename:
            continue
        safe_asset = sanitize_upload_name(asset.filename)
        if Path(safe_asset).suffix.lower() in {".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg"}:
            asset.save(str(upload_dir / safe_asset))

    results = []
    for md_file in md_files:
        if not md_file.filename or not md_file.filename.lower().endswith((".md", ".markdown")):
            continue
        safe_name = sanitize_upload_name(md_file.filename)
        source_path = upload_dir / safe_name
        md_file.save(str(source_path))
        item = {"original_name": safe_name, "status": "done", "outputs": [], "error": None}
        try:
            content = source_path.read_text(encoding="utf-8-sig")
            title = source_path.stem
            if export_pdf:
                pdf_name = f"{title}.pdf"
                markdown_to_pdf(content, output_dir / pdf_name, upload_dir, title)
                item["outputs"].append({"format": "PDF", "filename": pdf_name})
            if export_word:
                docx_name = f"{title}.docx"
                markdown_to_docx(content, output_dir / docx_name, upload_dir, title)
                item["outputs"].append({"format": "Word", "filename": docx_name})
        except Exception as exc:
            item["status"] = "error"
            item["error"] = str(exc)
        results.append(item)

    if not results:
        raise ValueError("No se encontraron archivos Markdown válidos")

    zip_path = output_dir / "exports_pdf_word.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in output_dir.iterdir():
            if path.suffix.lower() in {".pdf", ".docx"}:
                archive.write(path, path.name)

    job = {
        "id": job_id,
        "status": "done",
        "type": "markdown_export",
        "created_at": datetime.now().isoformat(),
        "finished_at": datetime.now().isoformat(),
        "total": len(results),
        "completed": len(results),
        "files": results,
    }
    with jobs_lock:
        jobs[job_id] = job
        persist_job_snapshot_locked(job_id)
    return job


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

    job_id = datetime.now().strftime("%Y%m%d_%H%M%S_") + uuid4().hex[:8]
    job_upload_dir = UPLOAD_DIR / job_id
    job_upload_dir.mkdir(parents=True, exist_ok=True)

    file_list = []
    used_names = set()
    for f in files:
        if f.filename and is_supported_convert_file(f.filename):
            safe_name = unique_upload_name(f.filename, used_names)
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

    ocr_dpi = max(72, min(ocr_dpi or 300, 600))
    ocr_mode = request.form.get("ocrMode", "auto")
    if ocr_mode not in {"auto", "force", "none"}:
        ocr_mode = "auto"
    options = {
        "extract_images": request.form.get("optImages") == "true",
        "exclude_headers": request.form.get("optHeaders") == "true",
        "ocr_mode": ocr_mode,
        "ocr_dpi": ocr_dpi,
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
        persist_job_snapshot_locked(job_id)

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

    job_id = datetime.now().strftime("%Y%m%d_%H%M%S_merge_") + uuid4().hex[:8]

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
        persist_job_snapshot_locked(job_id)

    return jsonify({"job_id": job_id, "total": len(md_files_data), "status": "done"})


@app.route("/export-md", methods=["POST"])
def export_md():
    """Convert standard Markdown to PDF and/or Word."""
    md_files = request.files.getlist("mds")
    asset_files = request.files.getlist("assets")
    if not md_files or all(not item.filename for item in md_files):
        return jsonify({"error": "Selecciona al menos un archivo Markdown"}), 400
    export_pdf = request.form.get("exportPdf", "true") == "true"
    export_word = request.form.get("exportWord", "true") == "true"
    if not export_pdf and not export_word:
        return jsonify({"error": "Selecciona PDF, Word o ambos formatos"}), 400
    try:
        job = export_markdown_job(md_files, asset_files, export_pdf, export_word)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    return jsonify(job)


@app.route("/download-exports/<job_id>")
def download_exports(job_id):
    zip_path = resolve_job_file(OUTPUT_DIR, job_id, "exports_pdf_word.zip")
    if zip_path is None or not zip_path.is_file():
        return jsonify({"error": "El paquete de exportación no existe"}), 404
    return send_file(str(zip_path), as_attachment=True)


@app.route("/status/<job_id>")
def status(job_id):
    """Get job status."""
    if resolve_job_directory(OUTPUT_DIR, job_id) is None:
        return jsonify({"error": "Job not found"}), 404
    with jobs_lock:
        job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify(job)


@app.route("/retry-errors/<job_id>", methods=["POST"])
def retry_errors(job_id):
    """Retry only failed files from a completed persisted job."""
    if resolve_job_directory(OUTPUT_DIR, job_id) is None:
        return jsonify({"error": "Job not found"}), 404
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return jsonify({"error": "Job not found"}), 404
        if job.get("status") == "running":
            return jsonify({"error": "El trabajo todavía está en ejecución"}), 409
        retry_count = 0
        for item in job.get("files", []):
            if item.get("status") == "error":
                item.update({
                    "status": "queued",
                    "error": None,
                    "time": None,
                    "size_md": None,
                    "output": None,
                })
                retry_count += 1
        if not retry_count:
            return jsonify({"error": "No hay archivos con error para reintentar"}), 400
        job["status"] = "running"
        job["finished_at"] = None
        job["completed"] = sum(
            item.get("status") == "done" for item in job.get("files", [])
        )
        persist_job_snapshot_locked(job_id)
    threading.Thread(target=run_batch_conversion, args=(job_id,), daemon=True).start()
    return jsonify({"job_id": job_id, "retried": retry_count, "status": "running"})


@app.route("/download/<job_id>/<filename>")
def download_file(job_id, filename):
    """Download a validated direct child of a job output directory."""
    file_path = resolve_job_file(OUTPUT_DIR, job_id, filename)
    if file_path is None or not file_path.is_file():
        return jsonify({"error": "File not found"}), 404
    return send_file(str(file_path), as_attachment=True)


@app.route("/download-all/<job_id>")
def download_all(job_id):
    """Download all converted files as ZIP."""
    zip_path = resolve_job_file(OUTPUT_DIR, job_id, "all_markdown.zip")
    if zip_path is None or not zip_path.is_file():
        return jsonify({"error": "ZIP not ready yet"}), 404
    return send_file(str(zip_path), as_attachment=True)


@app.route("/download-combined/<job_id>")
def download_combined(job_id):
    """Download all converted files as a single combined markdown."""
    combined_name = "all_combined.md"

    combined_path = resolve_job_file(OUTPUT_DIR, job_id, combined_name)
    if combined_path is None or not combined_path.is_file():
        return jsonify({"error": "Combined file not ready yet"}), 404
    return send_file(str(combined_path), as_attachment=True, download_name=combined_name)


@app.route("/download-merged/<job_id>")
def download_merged(job_id):
    """Download merged markdown file."""
    merged_path = resolve_job_file(OUTPUT_DIR, job_id, "merged_markdown.md")
    if merged_path is None or not merged_path.is_file():
        return jsonify({"error": "Merged file not found"}), 404
    return send_file(str(merged_path), as_attachment=True, download_name="merged_markdown.md")


@app.route("/preview/<job_id>/<filename>")
def preview(job_id, filename):
    """Get preview content from a validated Markdown output file."""
    file_path = resolve_job_file(OUTPUT_DIR, job_id, filename)
    if file_path is None or not file_path.is_file() or file_path.suffix.lower() != ".md":
        return jsonify({"error": "File not found"}), 404
    content = file_path.read_text(encoding="utf-8")
    if len(content) > 50000:
        content = content[:50000] + "\n\n... [truncated for preview] ..."
    return jsonify({"content": content, "filename": file_path.name})

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
        <button class="tab-btn" data-tab="export-panel" onclick="switchTab('export-panel')">
            📤 MD → PDF / Word
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
                <label><input type="checkbox" id="optImages"> Extraer imágenes embebidas (no recomendado para escaneados)</label>
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
                </div>
                <p style="font-size: 0.72rem; color: var(--text-dim); margin-top: 0.35rem; line-height: 1.4;">
                    <strong>Auto</strong> evalúa cada página y activa OCR solo en las páginas escaneadas. Aumentar DPI mejora la precisión del OCR, pero aumenta el tiempo.
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
            <button class="btn btn-secondary" id="retryErrorsBtn" onclick="retryErrors()" style="display:none; flex: 1; padding: 1rem; border-radius: 12px; font-size: 1.05rem;">
                ↻ <strong>Reintentar errores</strong><br>
                <span style="font-size: 0.75rem; opacity: 0.8;">Solo vuelve a procesar los archivos fallidos</span>
            </button>
        </div>

    </div>

    <!-- Markdown Export Panel -->
    <div class="tab-content" id="export-panel">
        <div class="dropzone" id="dropzoneExport" onclick="document.getElementById('fileInputExport').click()">
            <span class="dropzone-icon">📤</span>
            <h3>Convierte Markdown a PDF o Word</h3>
            <p>Admite Markdown estándar · agrega imágenes referenciadas si las necesitas</p>
        </div>
        <input type="file" id="fileInputExport" accept=".md,.markdown,.png,.jpg,.jpeg,.gif,.webp,.svg" multiple>

        <div class="options-panel active" style="margin-top: 1rem;">
            <h4>Formatos de salida</h4>
            <div class="option-row">
                <label><input type="checkbox" id="exportPdf" checked> PDF maquetado y buscable</label>
            </div>
            <div class="option-row">
                <label><input type="checkbox" id="exportWord" checked> Word editable (.docx)</label>
            </div>
        </div>

        <div class="file-queue" id="fileQueueExport"></div>
        <div class="actions" id="actionsBarExport" style="display: none;">
            <button class="btn btn-primary" id="exportBtn" onclick="startExport()">✨ Convertir Markdown</button>
            <button class="btn btn-secondary" onclick="clearAllExport()">✕ Limpiar</button>
            <span id="fileCountExport" style="color: var(--text-dim); font-size: 0.85rem; font-family: 'JetBrains Mono', monospace;"></span>
        </div>

        <div class="results-actions" id="exportResults">
            <div style="width: 100%;">
                <h3 style="font-size: 1.1rem; margin-bottom: 1rem; color: var(--accent);">Exportación completada</h3>
                <div class="file-queue" id="exportResultFiles"></div>
            </div>
            <button class="btn btn-download" onclick="downloadExports()" style="flex: 1; padding: 1rem;">
                📦 <strong>Descargar todas las exportaciones</strong><br>
                <span style="font-size: 0.75rem; opacity: 0.8;">PDF y Word en un ZIP</span>
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
let pendingExportFiles = [];
let currentJobId = null;
let pollInterval = null;

const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const fileQueue = document.getElementById('fileQueue');

const dropzoneMd = document.getElementById('dropzoneMd');
const fileInputMd = document.getElementById('fileInputMd');
const fileQueueMd = document.getElementById('fileQueueMd');
const dropzoneExport = document.getElementById('dropzoneExport');
const fileInputExport = document.getElementById('fileInputExport');
const fileQueueExport = document.getElementById('fileQueueExport');

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

// ── Drag & Drop Markdown Export ──
['dragenter', 'dragover'].forEach(function(eventName) {
    dropzoneExport.addEventListener(eventName, function(event) {
        event.preventDefault();
        dropzoneExport.classList.add('dragover');
    });
});
['dragleave', 'drop'].forEach(function(eventName) {
    dropzoneExport.addEventListener(eventName, function(event) {
        event.preventDefault();
        dropzoneExport.classList.remove('dragover');
    });
});
dropzoneExport.addEventListener('drop', function(event) {
    addExportFiles(Array.from(event.dataTransfer.files).filter(isMarkdownExportFile));
});
fileInputExport.addEventListener('change', function() {
    addExportFiles(Array.from(fileInputExport.files));
    fileInputExport.value = '';
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

function isMarkdownExportFile(file) {
    return /\.(md|markdown|png|jpe?g|gif|webp|svg)$/i.test(file.name);
}

function addExportFiles(files) {
    files.filter(isMarkdownExportFile).forEach(function(file) {
        if (!pendingExportFiles.find(function(current) { return current.name === file.name && current.size === file.size; })) {
            pendingExportFiles.push(file);
        }
    });
    renderExportQueue();
}

function removeExportFile(index) {
    pendingExportFiles.splice(index, 1);
    renderExportQueue();
}

function clearAllExport() {
    pendingExportFiles = [];
    window.exportJobId = null;
    renderExportQueue();
    document.getElementById('exportResults').classList.remove('active');
}

function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}
function escapeHtml(value) {
    return String(value ?? '')
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#039;');
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
            <span class="name" title="${escapeHtml(f.name)}">${escapeHtml(f.name)}</span>
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
            <span class="name" title="${escapeHtml(f.name)}">${escapeHtml(f.name)}</span>
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

function renderExportQueue() {
    fileQueueExport.innerHTML = '';
    pendingExportFiles.forEach(function(file, index) {
        const div = document.createElement('div');
        const isMd = /\.(md|markdown)$/i.test(file.name);
        div.className = 'file-item';
        div.innerHTML =
            '<span class="icon">' + (isMd ? '📝' : '🖼️') + '</span>' +
            '<span class="name" title="' + escapeHtml(file.name) + '">' + escapeHtml(file.name) + '</span>' +
            '<span class="size">' + formatSize(file.size) + '</span>' +
            '<button class="remove-btn" onclick="removeExportFile(' + index + ')" title="Quitar">×</button>';
        fileQueueExport.appendChild(div);
    });
    const mdCount = pendingExportFiles.filter(function(file) { return /\.(md|markdown)$/i.test(file.name); }).length;
    document.getElementById('actionsBarExport').style.display = mdCount ? 'flex' : 'none';
    document.getElementById('fileCountExport').textContent =
        mdCount + ' Markdown · ' + (pendingExportFiles.length - mdCount) + ' recurso(s)';
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
        pollInterval = setInterval(pollStatus, 2500);

    } catch (err) {
        alert('Error de conexión: ' + err.message);
        btn.disabled = false;
        btn.innerHTML = '⚡ Convertir todo';
    }
}

async function startExport() {
    const markdownFiles = pendingExportFiles.filter(function(file) { return /\.(md|markdown)$/i.test(file.name); });
    if (!markdownFiles.length) return;
    const exportPdf = document.getElementById('exportPdf').checked;
    const exportWord = document.getElementById('exportWord').checked;
    if (!exportPdf && !exportWord) {
        alert('Selecciona PDF, Word o ambos formatos.');
        return;
    }
    const btn = document.getElementById('exportBtn');
    btn.disabled = true;
    btn.textContent = '⏳ Formateando...';
    const formData = new FormData();
    markdownFiles.forEach(function(file) { formData.append('mds', file); });
    pendingExportFiles.filter(function(file) { return !/\.(md|markdown)$/i.test(file.name); })
        .forEach(function(file) { formData.append('assets', file); });
    formData.append('exportPdf', exportPdf);
    formData.append('exportWord', exportWord);
    try {
        const response = await fetch('/export-md', { method: 'POST', body: formData });
        const data = await response.json();
        if (!response.ok || data.error) throw new Error(data.error || 'No fue posible exportar');
        window.exportJobId = data.id;
        const resultBox = document.getElementById('exportResultFiles');
        resultBox.innerHTML = '';
        data.files.forEach(function(file) {
            const div = document.createElement('div');
            div.className = 'file-item';
            const links = (file.outputs || []).map(function(output) {
                return '<a class="btn btn-secondary" style="padding:0.3rem 0.65rem;text-decoration:none;" href="/download/' +
                    data.id + '/' + encodeURIComponent(output.filename) + '">↓ ' + output.format + '</a>';
            }).join('');
            div.innerHTML = '<span class="icon">📄</span><span class="name">' + escapeHtml(file.original_name) +
                '</span><span class="status-badge status-' + file.status + '">' +
                (file.status === 'done' ? '✓ Listo' : '✗ Error') +
                '</span><span style="display:flex;gap:0.35rem;">' + links + '</span>';
            resultBox.appendChild(div);
        });
        document.getElementById('exportResults').classList.add('active');
        btn.textContent = '✓ Completado';
    } catch (error) {
        alert('Error: ' + error.message);
        btn.textContent = '✨ Convertir Markdown';
    } finally {
        btn.disabled = false;
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
            actions =
                '<a href="/download/' + job.id + '/' + encodeURIComponent(f.output) +
                '" class="btn btn-secondary" style="padding:0.25rem 0.6rem;font-size:0.72rem;text-decoration:none;">↓ MD</a>' +
                '<button class="btn btn-secondary" style="padding:0.25rem 0.6rem;font-size:0.72rem;" onclick="showPreview(' +
                "'" + job.id + "','" + f.output + "'" + ')">👁</button>';
        }
        if (f.status === 'error') {
            actions = '<span style="color:var(--red);font-size:0.72rem;" title="' +
                escapeHtml(f.error || '') + '">ver error</span>';
        }

        let timeInfo = f.time ? `${f.time}s` : '';

        div.innerHTML = `
            <span class="icon">📄</span>
            <span class="name" title="${escapeHtml(f.original_name)}">${escapeHtml(f.original_name)}</span>
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
        const retryButton = document.getElementById('retryErrorsBtn');
        retryButton.style.display = errors > 0 && job.status === 'done' ? 'inline-flex' : 'none';
        const totalTime = job.files.reduce((s, f) => s + (f.time || 0), 0);
        document.getElementById('statDone').textContent = done;
        document.getElementById('statErrors').textContent = errors;
        document.getElementById('statTime').textContent = totalTime.toFixed(1) + 's';
    }
}

// ── Download All ──
async function retryErrors() {
    if (!currentJobId) return;
    const button = document.getElementById('retryErrorsBtn');
    button.disabled = true;
    try {
        const response = await fetch(`/retry-errors/${currentJobId}`, { method: 'POST' });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || 'No fue posible reintentar');
        document.getElementById('resultsActions').classList.remove('active');
        document.getElementById('convertBtn').innerHTML = '⏳ Convirtiendo...';
        pollInterval = setInterval(pollStatus, 2500);
    } catch (error) {
        alert('Error: ' + error.message);
    } finally {
        button.disabled = false;
    }
}

function downloadAll() {
    if (currentJobId) {
        window.location.href = `/download-all/${currentJobId}`;
    }
}

function downloadCombined() {
    if (currentJobId) {
        window.location.href = '/download-combined/' + currentJobId;
    }
}

function downloadExports() {
    if (window.exportJobId) {
        window.location.href = '/download-exports/' + window.exportJobId;
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
    for resumable_job_id in load_job_snapshots():
        threading.Thread(
            target=run_batch_conversion,
            args=(resumable_job_id,),
            daemon=True,
        ).start()
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 5000
    print(f"""
+----------------------------------------------------+
|   PDF -> Markdown  ·  Batch Converter             |
|   http://localhost:{port}                          |
|   Ctrl+C para detener                            |
+----------------------------------------------------+
    """)
    app.run(host="127.0.0.1", port=port, debug=False, threaded=True)








