#!/usr/bin/env python3
"""
PDF2MD Batch Converter
=====================
Local web application for batch converting PDF, Word and RTF files to Markdown.
Uses pymupdf4llm for high-quality extraction with table, image, and header support.

Usage:
    python app.py
    Then open http://localhost:5000 in your browser.
"""

import os
import platform
import sys
import json
import re
import time
import zipfile
import shutil
import threading
import copy
import logging
import subprocess
from collections import Counter
from pathlib import Path
from datetime import datetime
from logging.handlers import RotatingFileHandler
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
from runtime_config import (
    APP_VERSION,
    DATA_DIR,
    DATA_RETENTION_DAYS,
    LOG_DIR,
    OUTPUT_DIR,
    UPLOAD_DIR,
    configure_tesseract,
    ensure_runtime_directories,
    resource_root,
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

try:
    from striprtf.striprtf import rtf_to_text
except ImportError:
    rtf_to_text = None

# ─── Configuration ───────────────────────────────────────────────────────────

app = Flask(__name__)
# Uploads are sent one file at a time by the desktop UI. Keep a generous
# per-request ceiling for unusually large legal records and for compatibility
# with older browser pages that still send a whole batch in one request.
MAX_UPLOAD_REQUEST_BYTES = 8 * 1024 * 1024 * 1024
MIN_FREE_SPACE_AFTER_UPLOAD = 256 * 1024 * 1024
app.config['MAX_CONTENT_LENGTH'] = MAX_UPLOAD_REQUEST_BYTES

ensure_runtime_directories()

LOG_FILE = LOG_DIR / "pdf2md.log"
log_handler = RotatingFileHandler(
    LOG_FILE, maxBytes=2 * 1024 * 1024, backupCount=3, encoding="utf-8"
)
log_handler.setFormatter(logging.Formatter(
    "%(asctime)s | %(levelname)s | %(threadName)s | %(message)s"
))
app.logger.setLevel(logging.INFO)
if not any(isinstance(handler, RotatingFileHandler) for handler in app.logger.handlers):
    app.logger.addHandler(log_handler)

WORD_EXTENSIONS = {".docx", ".docm"}
RTF_EXTENSIONS = {".rtf"}
SUPPORTED_CONVERT_EXTENSIONS = {".pdf", *WORD_EXTENSIONS, *RTF_EXTENSIONS}


TESSDATA_DIR = configure_tesseract()
ASSET_DIR = resource_root() / "assets"
if TESSDATA_DIR:
    app.logger.info("OCR configurado con tessdata en %s", TESSDATA_DIR)
else:
    app.logger.warning("No se encontró tessdata en español e inglés")

# Track conversion jobs
jobs = {}
jobs_lock = threading.Lock()
ocr_lock = threading.Lock()
OCR_MAX_ATTEMPTS = 3
OCR_RETRY_DELAY_SECONDS = 0.5
initialization_lock = threading.Lock()
application_initialized = False


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


def cleanup_expired_jobs(retention_days: int = DATA_RETENTION_DAYS) -> int:
    """Remove old internal jobs while preserving active work and user downloads."""
    cutoff = time.time() - max(retention_days, 1) * 86400
    removed = 0
    for root in (UPLOAD_DIR, OUTPUT_DIR):
        for child in root.iterdir():
            if not child.is_dir():
                continue
            try:
                if child.stat().st_mtime >= cutoff:
                    continue
                snapshot = OUTPUT_DIR / child.name / "job.json"
                if snapshot.is_file():
                    job = json.loads(snapshot.read_text(encoding="utf-8"))
                    if job.get("status") in {"running", "packaging"}:
                        continue
                shutil.rmtree(child)
                removed += 1
            except (OSError, ValueError, TypeError) as exc:
                app.logger.warning("No se pudo limpiar %s: %s", child, exc)
    return removed


def initialize_application() -> None:
    """Run safe startup maintenance and resume interrupted conversions once."""
    global application_initialized
    with initialization_lock:
        if application_initialized:
            return
        removed = cleanup_expired_jobs()
        if removed:
            app.logger.info("Se eliminaron %s carpetas internas antiguas", removed)
        for resumable_job_id in load_job_snapshots():
            threading.Thread(
                target=run_batch_conversion,
                args=(resumable_job_id,),
                daemon=True,
                name=f"resume-{resumable_job_id}",
            ).start()
        application_initialized = True
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


def normalize_marginal_text(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip().casefold()


def detect_repeated_marginal_text(doc) -> set[str]:
    """Find actual repeated headers/footers instead of cropping page content."""
    if len(doc) < 2:
        return set()
    occurrences = Counter()
    for page in doc:
        page_height = float(page.rect.height or 1)
        seen_on_page = set()
        for block in page.get_text("blocks", sort=True):
            if len(block) < 7 or int(block[6]) != 0:
                continue
            y0, y1 = float(block[1]), float(block[3])
            if y1 > page_height * 0.1 and y0 < page_height * 0.9:
                continue
            normalized = normalize_marginal_text(str(block[4]))
            if 3 <= len(normalized) <= 300:
                seen_on_page.add(normalized)
        occurrences.update(seen_on_page)
    threshold = max(2, (len(doc) + 1) // 2)
    return {text for text, count in occurrences.items() if count >= threshold}


def native_page_to_markdown(page, repeated_marginal_text: set[str] | None = None) -> str:
    """Extract selectable text without invoking the CPU-heavy layout model.

    Digital legal PDFs already contain positioned text. Re-running the neural
    layout engine on every such page can be slower than OCR on modest CPUs.
    This path preserves blocks, line order and simple headings while keeping
    the full native text available for search and citation.
    """
    flags = fitz.TEXTFLAGS_DICT | fitz.TEXT_DEHYPHENATE
    page_dict = page.get_text("dict", flags=flags, sort=True)
    blocks = [block for block in page_dict.get("blocks", []) if block.get("type") == 0]
    font_sizes = Counter()
    for block in blocks:
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = str(span.get("text") or "")
                if text.strip():
                    font_sizes[round(float(span.get("size") or 0), 1)] += len(text.strip())
    body_size = font_sizes.most_common(1)[0][0] if font_sizes else 10.0
    page_height = float(page.rect.height or 1)
    markdown_blocks = []

    for block in blocks:
        bbox = block.get("bbox") or (0, 0, 0, 0)
        lines = []
        block_sizes = []
        bold_characters = 0
        total_characters = 0
        for line in block.get("lines", []):
            line_parts = []
            for span in line.get("spans", []):
                span_text = str(span.get("text") or "")
                line_parts.append(span_text)
                characters = len(span_text.strip())
                total_characters += characters
                block_sizes.append(float(span.get("size") or body_size))
                font_name = str(span.get("font") or "").casefold()
                if "bold" in font_name or "black" in font_name or "semibold" in font_name:
                    bold_characters += characters
            line_text = "".join(line_parts).strip()
            if line_text:
                lines.append(line_text)
        if not lines:
            continue

        block_text = "\n".join(lines).strip()
        is_marginal = float(bbox[3]) <= page_height * 0.1 or float(bbox[1]) >= page_height * 0.9
        if (
            repeated_marginal_text
            and is_marginal
            and normalize_marginal_text(block_text) in repeated_marginal_text
        ):
            continue
        maximum_size = max(block_sizes or [body_size])
        mostly_bold = total_characters > 0 and bold_characters / total_characters >= 0.6
        letters = re.sub(r"[^A-Za-zÁÉÍÓÚÜÑáéíóúüñ]", "", block_text)
        mostly_upper = bool(letters) and sum(char.isupper() for char in letters) / len(letters) >= 0.8
        is_heading = (
            len(block_text) <= 180
            and len(lines) <= 3
            and (maximum_size >= body_size * 1.22 or (mostly_bold and mostly_upper))
        )
        if is_heading:
            level = 2 if maximum_size >= body_size * 1.55 else 3
            markdown_blocks.append(f"{'#' * level} {block_text.replace(chr(10), ' ')}")
        else:
            markdown_blocks.append(block_text)

    return clean_page_markdown("\n\n".join(markdown_blocks))
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


def rtf_document_to_markdown(rtf_content: bytes) -> str:
    """Extract readable text from an RTF document as portable Markdown."""
    if rtf_to_text is None:
        raise RuntimeError("striprtf is not installed. Run: pip install striprtf")

    # RTF control words are ASCII. Latin-1 preserves every byte one-to-one so
    # striprtf can honor the document's own ANSI code-page declaration.
    source = rtf_content.decode("latin-1")
    if not source.lstrip().startswith("{\\rtf"):
        raise ValueError("El archivo no contiene un documento RTF válido.")

    plain_text = rtf_to_text(source, errors="replace")
    plain_text = plain_text.replace("\r\n", "\n").replace("\r", "\n")
    markdown_lines = []
    for raw_line in plain_text.split("\n"):
        line = raw_line.replace("\t", "    ").strip()
        if line.startswith(("• ", "· ", "◦ ")):
            line = f"- {line[2:].strip()}"
        if line:
            markdown_lines.append(line)

    markdown_text = "\n\n".join(markdown_lines).strip()
    if not markdown_text:
        raise ValueError("El documento RTF no produjo texto reconocible.")
    return markdown_text + "\n"


def is_image_only_markdown(markdown_text: str) -> bool:
    """Detect output that contains image references but no useful text."""
    stripped = (markdown_text or "").strip()
    if not stripped:
        return True
    has_images = bool(re.search(r"!\[[^\]]*\]\([^)]*\)|!\[\[[^\]]+\]\]", stripped))
    return has_images and markdown_semantic_word_count(stripped) < 10


def update_file_progress(job_id: str, file_index: int, **changes) -> None:
    """Publish lightweight per-file progress without rewriting snapshots per page."""
    with jobs_lock:
        job = jobs.get(job_id)
        if not job or file_index >= len(job.get("files", [])):
            return
        job["files"][file_index].update(changes)
        job["updated_at"] = datetime.now().isoformat()


def friendly_conversion_error(exc: Exception) -> str:
    """Translate common technical failures into actionable Spanish messages."""
    detail = str(exc).strip()
    lowered = detail.casefold()
    if any(term in lowered for term in ("password", "encrypted", "authenticate", "cifrad")):
        return "El documento está protegido con contraseña. Abre una copia sin protección e inténtalo de nuevo."
    if any(term in lowered for term in ("xref", "cannot open", "damaged", "broken", "corrupt")):
        return "El PDF parece estar dañado o incompleto. Intenta abrirlo y guardarlo nuevamente como PDF."
    if any(term in lowered for term in ("tesseract", "tessdata", "ocr", "leptonica")):
        return "No fue posible reconocer el texto de una o más páginas escaneadas. Revisa la calidad del PDF y vuelve a intentarlo."
    if "no produjo texto" in lowered or "sin texto reconocible" in lowered:
        return "El documento no contiene texto reconocible. Prueba la opción Forzar OCR si se trata de un escaneo."
    return "No fue posible convertir este archivo. Puedes reintentarlo o consultar el detalle técnico en Acerca de y ayuda."


def conversion_options_from(data) -> dict:
    """Normalize conversion options from either a form or a JSON object."""
    try:
        ocr_dpi = int(data.get("ocrDpi") or 300)
    except (TypeError, ValueError):
        ocr_dpi = 300
    ocr_mode = str(data.get("ocrMode") or "auto")
    if ocr_mode not in {"auto", "force", "none"}:
        ocr_mode = "auto"

    def enabled(value) -> bool:
        return value is True or str(value).casefold() == "true"

    return {
        "extract_images": enabled(data.get("optImages")),
        "exclude_headers": enabled(data.get("optHeaders")),
        "ocr_mode": ocr_mode,
        "ocr_dpi": max(72, min(ocr_dpi, 600)),
    }


def create_conversion_job(options: dict, status_value: str = "uploading") -> str:
    """Create a persisted job before receiving its files."""
    job_id = datetime.now().strftime("%Y%m%d_%H%M%S_") + uuid4().hex[:8]
    (UPLOAD_DIR / job_id).mkdir(parents=True, exist_ok=False)
    with jobs_lock:
        jobs[job_id] = {
            "id": job_id,
            "status": status_value,
            "created_at": datetime.now().isoformat(),
            "started_timestamp": None,
            "finished_at": None,
            "total": 0,
            "completed": 0,
            "files": [],
            "options": options,
        }
        persist_job_snapshot_locked(job_id)
    return job_id


def save_file_in_job(job_id: str, uploaded_file) -> dict:
    """Save one uploaded document atomically and append it to an upload job."""
    if not uploaded_file or not uploaded_file.filename:
        raise ValueError("No se recibió ningún archivo.")
    if not is_supported_convert_file(uploaded_file.filename):
        raise ValueError("El archivo no es PDF, Word ni RTF compatible.")

    with jobs_lock:
        job = jobs.get(job_id)
        if not job or job.get("status") != "uploading":
            raise ValueError("La carga ya no está disponible. Inicia la conversión nuevamente.")
        used_names = {
            str(item.get("original_name") or "").casefold()
            for item in job.get("files", [])
        }
        safe_name = unique_upload_name(uploaded_file.filename, used_names)

    job_upload_dir = resolve_job_directory(UPLOAD_DIR, job_id)
    if job_upload_dir is None:
        raise ValueError("El identificador de carga no es válido.")
    job_upload_dir.mkdir(parents=True, exist_ok=True)
    free_space = shutil.disk_usage(DATA_DIR).free
    announced_size = max(int(request.content_length or 0), 0)
    if free_space - announced_size < MIN_FREE_SPACE_AFTER_UPLOAD:
        raise OSError(
            "No hay espacio suficiente en el disco para copiar este archivo. "
            "Libera espacio y vuelve a intentarlo."
        )

    save_path = job_upload_dir / safe_name
    partial_path = save_path.with_suffix(save_path.suffix + ".part")
    try:
        uploaded_file.save(str(partial_path))
        os.replace(partial_path, save_path)
    finally:
        if partial_path.exists():
            try:
                partial_path.unlink()
            except OSError:
                pass

    file_info = {
        "original_name": safe_name,
        "size": save_path.stat().st_size,
        "status": "queued",
        "output": None,
        "error": None,
        "time": None,
        "size_md": None,
        "type": save_path.suffix.lower().lstrip("."),
        "current_page": 0,
        "total_pages": None,
        "ocr_active": False,
        "ocr_pages_completed": 0,
        "extraction_mode": "queued",
    }
    with jobs_lock:
        job = jobs.get(job_id)
        if not job or job.get("status") != "uploading":
            try:
                save_path.unlink()
            except OSError:
                pass
            raise ValueError("La carga fue cancelada antes de terminar.")
        job["files"].append(file_info)
        job["total"] = len(job["files"])
        persist_job_snapshot_locked(job_id)
    return file_info

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

def pdf_to_markdown_page_by_page(
    pdf_path: Path,
    output_path: Path,
    options: dict,
    progress_callback=None,
) -> str:
    """Convert every page independently, applying OCR only where required."""
    ocr_mode = options.get("ocr_mode", "auto")
    dpi = max(72, min(int(options.get("ocr_dpi") or 300), 600))
    write_images = bool(options.get("extract_images", False))
    page_blocks = []
    ocr_pages_completed = 0

    with fitz.open(str(pdf_path)) as doc:
        total_pages = len(doc)
        repeated_marginal_text = (
            detect_repeated_marginal_text(doc)
            if options.get("exclude_headers", False)
            else set()
        )
        if progress_callback:
            progress_callback(current_page=0, total_pages=total_pages, ocr_active=False)
        for page_index, page in enumerate(doc):
            page_number = page_index + 1
            native_text = (page.get_text("text") or "").strip()
            page_has_images = bool(page.get_images(full=True))
            has_usable_native_text = page_has_usable_text(page)
            should_ocr = ocr_mode == "force" or (
                ocr_mode == "auto"
                and not has_usable_native_text
                and page_has_images
            )
            page_extraction_mode = "ocr" if should_ocr else "layout"

            page_markdown = ""
            if not should_ocr and has_usable_native_text and not write_images:
                page_extraction_mode = "digital_fast"
                if progress_callback:
                    progress_callback(
                        current_page=page_number,
                        total_pages=total_pages,
                        ocr_active=False,
                        extraction_mode="digital_fast",
                        ocr_pages_completed=ocr_pages_completed,
                    )
                page_markdown = native_page_to_markdown(
                    page,
                    repeated_marginal_text=repeated_marginal_text,
                )
                # Native text is the authoritative fallback if block formatting
                # ever omits content from an unusual embedded font.
                native_words = markdown_semantic_word_count(native_text)
                if markdown_semantic_word_count(page_markdown) < max(5, int(native_words * 0.75)):
                    page_markdown = native_text
            elif not should_ocr:
                page_extraction_mode = "layout"
                if progress_callback:
                    progress_callback(
                        current_page=page_number,
                        total_pages=total_pages,
                        ocr_active=False,
                        extraction_mode="layout",
                        ocr_pages_completed=ocr_pages_completed,
                    )
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
                page_extraction_mode = "ocr"
                if progress_callback:
                    progress_callback(
                        current_page=page_number,
                        total_pages=total_pages,
                        ocr_active=True,
                        extraction_mode="ocr",
                        ocr_pages_completed=ocr_pages_completed,
                    )
                page_markdown = clean_page_markdown(
                    ocr_page_to_markdown(page, page_number, dpi)
                )
                if not page_markdown:
                    if page_has_images and not options.get("allow_empty_ocr_pages", False):
                        raise RuntimeError(
                            f"OCR no devolvió texto en la página {page_number}. "
                            "Verifica la calidad del escaneo."
                        )
                    page_markdown = "_Página sin texto reconocible._"
                ocr_pages_completed += 1

            if not page_markdown:
                page_markdown = "_Página sin texto reconocible._"
            page_blocks.append(f"## Página {page_number}\n\n{page_markdown}".strip())
            if progress_callback:
                progress_callback(
                    current_page=page_number,
                    total_pages=total_pages,
                    ocr_active=False,
                    extraction_mode=page_extraction_mode,
                    ocr_pages_completed=ocr_pages_completed,
                )

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
        md_text = pdf_to_markdown_page_by_page(
            pdf_path,
            output_path,
            options,
            progress_callback=lambda **changes: update_file_progress(
                job_id, file_index, **changes
            ),
        )
        output_name = write_markdown_output(md_text, output_path)
        elapsed = round(time.time() - start_time, 2)

        with jobs_lock:
            jobs[job_id]["files"][file_index]["status"] = "done"
            jobs[job_id]["files"][file_index]["output"] = output_name
            jobs[job_id]["files"][file_index]["size_md"] = len(md_text)
            jobs[job_id]["files"][file_index]["time"] = elapsed
            jobs[job_id]["files"][file_index]["ocr_active"] = False
            jobs[job_id]["completed"] += 1
            persist_job_snapshot_locked(job_id)

    except Exception as e:
        app.logger.exception("Error convirtiendo PDF %s", pdf_path.name)
        with jobs_lock:
            jobs[job_id]["files"][file_index]["status"] = "error"
            jobs[job_id]["files"][file_index]["error"] = friendly_conversion_error(e)
            jobs[job_id]["files"][file_index]["technical_error"] = str(e)
            jobs[job_id]["files"][file_index]["ocr_active"] = False
            jobs[job_id]["files"][file_index]["time"] = round(time.time() - start_time, 2)
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
        app.logger.exception("Error convirtiendo Word %s", word_path.name)
        with jobs_lock:
            jobs[job_id]["files"][file_index]["status"] = "error"
            jobs[job_id]["files"][file_index]["error"] = friendly_conversion_error(e)
            jobs[job_id]["files"][file_index]["technical_error"] = str(e)
            jobs[job_id]["files"][file_index]["time"] = round(time.time() - start_time, 2)
            jobs[job_id]["completed"] += 1
            persist_job_snapshot_locked(job_id)


def convert_rtf_to_md(rtf_path: Path, output_path: Path, job_id: str, file_index: int):
    """Convert an RTF file to portable Markdown."""
    start_time = time.time()
    try:
        md_text = rtf_document_to_markdown(rtf_path.read_bytes())
        output_name = write_markdown_output(md_text, output_path)
        elapsed = round(time.time() - start_time, 2)
        with jobs_lock:
            jobs[job_id]["files"][file_index]["status"] = "done"
            jobs[job_id]["files"][file_index]["output"] = output_name
            jobs[job_id]["files"][file_index]["size_md"] = len(md_text)
            jobs[job_id]["files"][file_index]["time"] = elapsed
            jobs[job_id]["completed"] += 1
            persist_job_snapshot_locked(job_id)
    except Exception as exc:
        app.logger.exception("Error convirtiendo RTF %s", rtf_path.name)
        with jobs_lock:
            jobs[job_id]["files"][file_index]["status"] = "error"
            jobs[job_id]["files"][file_index]["error"] = friendly_conversion_error(exc)
            jobs[job_id]["files"][file_index]["technical_error"] = str(exc)
            jobs[job_id]["files"][file_index]["time"] = round(time.time() - start_time, 2)
            jobs[job_id]["completed"] += 1
            persist_job_snapshot_locked(job_id)


def convert_file_to_md(input_path: Path, output_path: Path, job_id: str, file_index: int, options: dict = None):
    """Convert one supported input file to Markdown based on its extension."""
    suffix = input_path.suffix.lower()
    if suffix == ".pdf":
        convert_pdf_to_md(input_path, output_path, job_id, file_index, options)
    elif suffix in WORD_EXTENSIONS:
        convert_word_to_md(input_path, output_path, job_id, file_index)
    elif suffix in RTF_EXTENSIONS:
        convert_rtf_to_md(input_path, output_path, job_id, file_index)
    else:
        with jobs_lock:
            jobs[job_id]["files"][file_index]["status"] = "error"
            jobs[job_id]["files"][file_index]["error"] = "Este tipo de archivo no es compatible."
            jobs[job_id]["files"][file_index]["technical_error"] = f"Unsupported file type: {suffix}"
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
        job["files"][file_index].update({
            "status": "converting",
            "started_at": datetime.now().isoformat(),
            "current_page": 0,
            "total_pages": None,
            "ocr_active": False,
            "ocr_pages_completed": 0,
            "extraction_mode": "queued",
        })
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
    return render_template_string(
        HTML_TEMPLATE,
        app_version=APP_VERSION,
        support_contact=os.environ.get(
            "PDF2MD_SUPPORT_CONTACT",
            "Consulta al docente o a la persona que te compartió la aplicación.",
        ),
        log_directory=str(LOG_DIR),
    )


@app.route("/health")
def health():
    return jsonify({
        "status": "ok",
        "version": APP_VERSION,
        "ocr_ready": TESSDATA_DIR is not None,
    })


@app.route("/assets/<path:filename>")
def packaged_asset(filename):
    return send_from_directory(str(ASSET_DIR), filename)


@app.route("/open-logs", methods=["POST"])
def open_logs():
    """Open the local diagnostics directory from the installed application."""
    try:
        if os.name == "nt":
            os.startfile(str(LOG_DIR))
        else:
            subprocess.Popen(["xdg-open", str(LOG_DIR)])
        return jsonify({"status": "ok", "path": str(LOG_DIR)})
    except OSError as exc:
        app.logger.exception("No se pudo abrir la carpeta de registros")
        return jsonify({
            "error": "No fue posible abrir la carpeta de diagnóstico.",
            "path": str(LOG_DIR),
            "detail": str(exc),
        }), 500


@app.route("/diagnostics")
def diagnostics():
    """Return copyable local diagnostics without sending anything externally."""
    disk = shutil.disk_usage(DATA_DIR)
    sections = [
        f"PDF2MD {APP_VERSION}",
        f"Sistema: {platform.platform()}",
        f"Python: {platform.python_version()}",
        f"Empaquetado: {bool(getattr(sys, 'frozen', False))}",
        f"Ejecutable: {sys.executable}",
        f"Datos: {DATA_DIR}",
        f"OCR: {TESSDATA_DIR or 'no disponible'}",
        f"Espacio libre: {disk.free} bytes",
    ]
    for name in ("launcher.log", "pdf2md.log", "crash.log"):
        log_path = LOG_DIR / name
        if not log_path.is_file():
            continue
        try:
            lines = log_path.read_text(encoding="utf-8", errors="replace").splitlines()[-80:]
            sections.append(f"\n--- {name} (últimas líneas) ---\n" + "\n".join(lines))
        except OSError as exc:
            sections.append(f"\n--- {name} ---\nNo se pudo leer: {exc}")
    return app.response_class("\n".join(sections), mimetype="text/plain; charset=utf-8")


@app.errorhandler(413)
def upload_too_large(_error):
    return jsonify({
        "error": (
            "La carga supera el límite de seguridad de 8 GB. "
            "Divide el lote en dos grupos y vuelve a intentarlo."
        )
    }), 413


@app.route("/upload/start", methods=["POST"])
def start_incremental_upload():
    """Create a job so the browser can upload large batches file by file."""
    options = conversion_options_from(request.get_json(silent=True) or {})
    try:
        job_id = create_conversion_job(options)
    except OSError as exc:
        app.logger.exception("No se pudo crear la carga incremental")
        return jsonify({"error": f"No fue posible preparar la carga: {exc}"}), 500
    return jsonify({"job_id": job_id, "status": "uploading"})


@app.route("/upload/file/<job_id>", methods=["POST"])
def upload_one_file(job_id):
    """Receive one file, avoiding a fragile multi-gigabyte HTTP request."""
    if resolve_job_directory(UPLOAD_DIR, job_id) is None:
        return jsonify({"error": "La carga ya no existe."}), 404
    uploaded_file = request.files.get("file")
    try:
        file_info = save_file_in_job(job_id, uploaded_file)
    except (ValueError, OSError) as exc:
        app.logger.warning("No se pudo cargar un archivo en %s: %s", job_id, exc)
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        app.logger.exception("Fallo inesperado al guardar un archivo en %s", job_id)
        return jsonify({
            "error": (
                "Windows no permitió guardar el archivo. Revisa el espacio disponible "
                "o la protección antivirus e inténtalo nuevamente."
            ),
            "detail": str(exc),
        }), 500
    return jsonify({"status": "uploaded", "file": file_info})


@app.route("/upload/finish/<job_id>", methods=["POST"])
def finish_incremental_upload(job_id):
    """Start conversion only after every file reached local storage."""
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return jsonify({"error": "La carga ya no existe."}), 404
        if job.get("status") != "uploading":
            return jsonify({"error": "Esta carga ya fue iniciada."}), 409
        if not job.get("files"):
            return jsonify({"error": "No se recibió ningún archivo compatible."}), 400
        job["status"] = "running"
        job["started_timestamp"] = time.time()
        job["total"] = len(job["files"])
        persist_job_snapshot_locked(job_id)
        total = job["total"]
    threading.Thread(
        target=run_batch_conversion,
        args=(job_id,),
        daemon=True,
        name=f"batch-{job_id}",
    ).start()
    return jsonify({"job_id": job_id, "total": total, "status": "running"})


@app.route("/upload/cancel/<job_id>", methods=["POST"])
def cancel_incremental_upload(job_id):
    """Remove only an incomplete upload created by the current browser action."""
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return jsonify({"status": "absent"})
        if job.get("status") != "uploading":
            return jsonify({"error": "La conversión ya comenzó y no se canceló."}), 409
        jobs.pop(job_id, None)
    for root in (UPLOAD_DIR, OUTPUT_DIR):
        target = resolve_job_directory(root, job_id)
        if target is not None and target.is_dir():
            try:
                shutil.rmtree(target)
            except OSError as exc:
                app.logger.warning("No se pudo limpiar la carga %s: %s", target, exc)
    return jsonify({"status": "cancelled"})


@app.route("/upload", methods=["POST"])
def upload():
    """Handle batch PDF/Word/RTF upload and start conversion."""
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
                "current_page": 0,
                "total_pages": None,
                "ocr_active": False,
                "ocr_pages_completed": 0,
                "extraction_mode": "queued",
            })

    if not file_list:
        return jsonify({"error": "No se encontraron archivos PDF, Word o RTF válidos (.pdf, .docx, .docm, .rtf)"}), 400

    options = conversion_options_from(request.form)

    with jobs_lock:
        jobs[job_id] = {
            "id": job_id,
            "status": "running",
            "created_at": datetime.now().isoformat(),
            "started_timestamp": time.time(),
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
        job = copy.deepcopy(jobs.get(job_id))
    if not job:
        return jsonify({"error": "Job not found"}), 404
    files = job.get("files", [])
    done = sum(item.get("status") == "done" for item in files)
    errors = sum(item.get("status") == "error" for item in files)
    active = [item for item in files if item.get("status") == "converting"]
    pending = sum(item.get("status") == "queued" for item in files)
    started = float(job.get("started_timestamp") or time.time())
    elapsed = max(0.0, time.time() - started)
    completed = done + errors
    partial = 0.0
    for item in active:
        total_pages = int(item.get("total_pages") or 0)
        current_page = int(item.get("current_page") or 0)
        if total_pages:
            partial += min(current_page / total_pages, 0.99)
    total = max(int(job.get("total") or 0), 1)
    progress_percent = min(100, round(((completed + partial) / total) * 100))
    eta = None
    if completed and job.get("status") != "done":
        eta = max(0, round((elapsed / completed) * (total - completed)))
    job["progress"] = {
        "done": done,
        "errors": errors,
        "pending": pending,
        "active": active,
        "ocr_active": any(item.get("ocr_active") for item in active),
        "elapsed_seconds": round(elapsed),
        "eta_seconds": eta,
        "percent": progress_percent,
    }
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
                    "technical_error": None,
                    "current_page": 0,
                    "total_pages": None,
                    "ocr_active": False,
                    "ocr_pages_completed": 0,
                    "extraction_mode": "queued",
                })
                retry_count += 1
        if not retry_count:
            return jsonify({"error": "No hay archivos con error para reintentar"}), 400
        job["status"] = "running"
        job["started_timestamp"] = time.time()
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
<title>PDF, Word y RTF → Markdown · Batch Converter</title>
<style>
@font-face { font-family: 'Outfit'; src: url('/assets/fonts/Outfit-Variable.ttf') format('truetype'); font-weight: 100 900; font-display: swap; }
@font-face { font-family: 'JetBrains Mono'; src: url('/assets/fonts/JetBrainsMono-Regular.ttf') format('truetype'); font-weight: 400; font-display: swap; }
@font-face { font-family: 'JetBrains Mono'; src: url('/assets/fonts/JetBrainsMono-Bold.ttf') format('truetype'); font-weight: 600 700; font-display: swap; }
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
.help-button {
    position: absolute;
    top: 1.25rem;
    right: 1.5rem;
    padding: 0.45rem 0.85rem;
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    color: var(--text-dim);
    cursor: pointer;
    font-family: 'Outfit', sans-serif;
    font-weight: 600;
}
.help-button:hover { color: var(--text); border-color: var(--accent); }
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
.progress-details {
    margin-top: 0.65rem;
    padding: 0.75rem 0.9rem;
    border: 1px solid var(--border);
    border-radius: 8px;
    background: var(--surface);
    color: var(--text-dim);
    font-size: 0.82rem;
    line-height: 1.55;
}
.progress-details strong { color: var(--text); }
.progress-details .ocr-notice { color: var(--yellow); }

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
.help-modal { max-width: 760px; }
.help-body {
    white-space: normal;
    font-family: 'Outfit', sans-serif;
    font-size: 0.93rem;
    line-height: 1.6;
}
.help-body h4 { color: var(--accent); margin: 1.1rem 0 0.35rem; }
.help-body h4:first-child { margin-top: 0; }
.help-body ul, .help-body ol { margin: 0.35rem 0 0.7rem 1.3rem; }
.help-body code {
    font-family: 'JetBrains Mono', monospace;
    color: var(--text);
    overflow-wrap: anywhere;
}
.privacy-note {
    margin-top: 1rem;
    padding: 0.8rem;
    border-left: 3px solid var(--green);
    background: rgba(52,211,153,0.08);
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
    .help-button { position: static; margin-top: 0.9rem; }
}
</style>
</head>
<body>

<div class="header">
    <h1>PDF, Word y RTF → Markdown</h1>
    <p>Conversión en lote · Local · Sin límites</p>
    <span class="badge">PDF · Word .docx/.docm · RTF · tablas · imágenes · multi-columna</span>
    <button class="help-button" onclick="openHelp()">? Acerca de y ayuda</button>
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
            <h3>Arrastra tus archivos PDF, Word o RTF aquí</h3>
            <p>o haz clic para seleccionar · acepta .pdf, .docx, .docm y .rtf</p>
        </div>
        <input type="file" id="fileInput" accept=".pdf,.docx,.docm,.rtf" multiple>

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
            <div class="progress-details" id="progressDetails">
                Preparando la conversión…
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
                <h3 style="font-size: 1.1rem; margin-bottom: 0.4rem; color: var(--accent);">✨ ¡Conversión completada!</h3>
                <p style="color: var(--text-dim); font-size: 0.88rem; line-height: 1.5;">Descarga los resultados o inicia otro lote desde aquí, sin actualizar la página.</p>
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
            <button class="btn btn-primary" id="convertMoreBtn" onclick="chooseMoreDocuments()" style="flex: 1 0 100%; padding: 1rem; border-radius: 12px; font-size: 1.05rem;" aria-label="Seleccionar más documentos para iniciar una nueva conversión">
                ＋ <strong>Convertir más documentos</strong><br>
                <span style="font-size: 0.75rem; opacity: 0.85;">Selecciona el siguiente lote sin recargar la página</span>
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

<!-- About and Help Modal -->
<div class="modal-overlay" id="helpModal">
    <div class="modal help-modal">
        <div class="modal-header">
            <h3>PDF2MD · Acerca de y ayuda</h3>
            <button class="modal-close" onclick="closeHelp()">×</button>
        </div>
        <div class="modal-body help-body">
            <h4>Uso básico</h4>
            <ol>
                <li>Selecciona o arrastra archivos PDF, DOCX, DOCM o RTF.</li>
                <li>Haz clic en <strong>Convertir todo</strong> y espera la confirmación.</li>
                <li>Descarga cada Markdown, el documento unido o el paquete ZIP.</li>
            </ol>

            <h4>Opciones de OCR</h4>
            <ul>
                <li><strong>Auto-detectar:</strong> usa extracción rápida en texto digital y reconoce únicamente las páginas escaneadas.</li>
                <li><strong>Forzar OCR:</strong> procesa todas las páginas como imágenes; úsalo en escaneos difíciles.</li>
                <li><strong>Desactivado:</strong> es más rápido, pero solo funciona con texto digital seleccionable.</li>
            </ul>

            <h4>Problemas frecuentes</h4>
            <ul>
                <li>Si un PDF pide contraseña, guarda primero una copia sin protección.</li>
                <li>Si el texto escaneado queda incompleto, prueba Forzar OCR a 300 DPI.</li>
                <li>Si un archivo falla, utiliza Reintentar errores; los demás continuarán normalmente.</li>
                <li>Las descargas suelen quedar en la carpeta Descargas configurada en tu navegador.</li>
            </ul>

            <div class="privacy-note">
                <strong>Privacidad:</strong> tus documentos se procesan únicamente en este computador.
                PDF2MD no los envía a internet y no recopila telemetría.
            </div>

            <h4>Soporte y diagnóstico</h4>
            <p>{{ support_contact }}</p>
            <p style="margin-top:0.45rem;">Versión <strong>{{ app_version }}</strong></p>
            <p style="margin-top:0.45rem;">Registros: <code>{{ log_directory }}</code></p>
            <button class="btn btn-secondary" style="margin-top:0.8rem;" onclick="openLogFolder()">
                📁 Abrir carpeta de diagnóstico
            </button>
            <button class="btn btn-secondary" style="margin-top:0.8rem;" onclick="copyDiagnostics()">
                📋 Copiar información de diagnóstico
            </button>
        </div>
    </div>
</div>

<script>
// ── State ──
let pendingFiles = [];
let pendingMdFiles = [];
let pendingExportFiles = [];
let currentJobId = null;
let pollInterval = null;
let conversionInProgress = false;
let consecutivePollFailures = 0;
let pollFailureReported = false;

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

function openHelp() {
    document.getElementById('helpModal').classList.add('active');
}

function closeHelp() {
    document.getElementById('helpModal').classList.remove('active');
}

async function openLogFolder() {
    try {
        const response = await fetch('/open-logs', { method: 'POST' });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || 'No fue posible abrir la carpeta.');
    } catch (error) {
        alert(error.message);
    }
}

async function copyDiagnostics() {
    try {
        const response = await fetch('/diagnostics', { cache: 'no-store' });
        if (!response.ok) throw new Error('No fue posible preparar el diagnóstico.');
        const diagnostic = await response.text();
        await navigator.clipboard.writeText(diagnostic);
        alert('La información de diagnóstico fue copiada. Ya puedes pegarla en tu mensaje de soporte.');
    } catch (error) {
        alert('No fue posible copiar automáticamente. Abre la carpeta de diagnóstico y comparte los archivos de registro.');
    }
}

document.getElementById('helpModal').addEventListener('click', event => {
    if (event.target.id === 'helpModal') closeHelp();
});

window.addEventListener('beforeunload', event => {
    if (!conversionInProgress) return;
    event.preventDefault();
    event.returnValue = 'Hay una conversión activa. Si cierras la aplicación, el proceso podría interrumpirse.';
});

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
    const convertibleFiles = files.filter(isConvertibleFile);
    if (convertibleFiles.length > 0 && isCompletedConversionVisible()) {
        resetConversionView();
        pendingFiles = [];
    }
    convertibleFiles.forEach(f => {
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
    resetConversionView();
    renderQueue();
}

function isCompletedConversionVisible() {
    return !conversionInProgress &&
        document.getElementById('resultsActions').classList.contains('active');
}

function resetConversionView() {
    currentJobId = null;
    if (pollInterval) {
        clearInterval(pollInterval);
        pollInterval = null;
    }
    conversionInProgress = false;
    consecutivePollFailures = 0;
    pollFailureReported = false;

    document.getElementById('progressContainer').classList.remove('active');
    document.getElementById('statsBar').classList.remove('active');
    document.getElementById('resultsActions').classList.remove('active');
    document.getElementById('retryErrorsBtn').style.display = 'none';

    document.getElementById('progressBar').style.width = '0%';
    document.getElementById('progressText').textContent = '0 / 0 archivos';
    document.getElementById('progressPercent').textContent = '0%';
    document.getElementById('progressDetails').textContent = 'Preparando la conversión…';

    const convertButton = document.getElementById('convertBtn');
    convertButton.disabled = false;
    convertButton.innerHTML = '⚡ Convertir todo';
}

function chooseMoreDocuments() {
    // The previous result remains visible if the user cancels the file picker.
    fileInput.value = '';
    fileInput.click();
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
    return /\.(pdf|docx|docm|rtf)$/i.test(file.name);
}

function fileIcon(filename) {
    return /\.(docx|docm|rtf)$/i.test(filename) ? '\u{1F4DD}' : '\u{1F4C4}';
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
async function fetchJson(url, options) {
    const response = await fetch(url, options);
    const raw = await response.text();
    let data = {};
    if (raw) {
        try {
            data = JSON.parse(raw);
        } catch (_error) {
            data = { error: raw.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim() };
        }
    }
    if (!response.ok) {
        throw new Error(data.error || `PDF2MD respondió con el código ${response.status}.`);
    }
    return data;
}

function showUploadProgress(file, index, total, loaded, fileTotal) {
    const completed = index;
    const fraction = fileTotal > 0 ? Math.min(loaded / fileTotal, 1) : 0;
    const percent = Math.round(((completed + fraction) / Math.max(total, 1)) * 100);
    document.getElementById('progressBar').style.width = percent + '%';
    document.getElementById('progressPercent').textContent = percent + '%';
    document.getElementById('progressText').textContent =
        `${completed} cargados · ${total - completed} pendientes · 0 errores`;
    document.getElementById('progressDetails').innerHTML =
        `<strong>Copiando ${escapeHtml(file.name)}</strong><br>` +
        `Archivo ${index + 1} de ${total} · ${formatSize(loaded)} de ${formatSize(fileTotal || file.size)}`;
}

function uploadSingleFile(jobId, file, index, total) {
    return new Promise((resolve, reject) => {
        const request = new XMLHttpRequest();
        request.open('POST', `/upload/file/${jobId}`);
        request.responseType = 'json';
        request.upload.addEventListener('progress', event => {
            showUploadProgress(file, index, total, event.loaded, event.total || file.size);
        });
        request.addEventListener('load', () => {
            const data = request.response || {};
            if (request.status >= 200 && request.status < 300) {
                resolve(data || {});
            } else {
                reject(new Error((data && data.error) || `No se pudo cargar ${file.name}.`));
            }
        });
        request.addEventListener('error', () => {
            const error = new Error('Se perdió la conexión con el servidor local durante la carga.');
            error.connectionLost = true;
            reject(error);
        });
        request.addEventListener('abort', () => reject(new Error('La carga fue cancelada.')));
        const formData = new FormData();
        formData.append('file', file);
        request.send(formData);
    });
}

async function explainConnectionFailure(error) {
    try {
        const health = await fetch('/health', { cache: 'no-store' });
        if (health.ok) {
            return `${error.message}\n\nPDF2MD continúa abierto. Revisa el espacio disponible y vuelve a intentarlo.`;
        }
    } catch (_healthError) {
        // The diagnostic below intentionally handles an unavailable server.
    }
    return 'El servidor local de PDF2MD se cerró o fue bloqueado por Windows. ' +
        'Abre PDF2MD nuevamente desde el escritorio. Si vuelve a ocurrir, comparte la carpeta ' +
        '%LOCALAPPDATA%\\PDF2MD\\logs con soporte.';
}

async function startConversion() {
    if (pendingFiles.length === 0) return;

    const btn = document.getElementById('convertBtn');
    btn.disabled = true;
    btn.innerHTML = '⏳ Subiendo...';
    conversionInProgress = true;
    consecutivePollFailures = 0;
    pollFailureReported = false;
    document.getElementById('progressContainer').classList.add('active');

    try {
        const started = await fetchJson('/upload/start', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                optImages: document.getElementById('optImages').checked,
                optHeaders: document.getElementById('optHeaders').checked,
                ocrMode: document.getElementById('ocrMode').value,
                ocrDpi: document.getElementById('ocrDpi').value,
            }),
        });
        currentJobId = started.job_id;

        for (let index = 0; index < pendingFiles.length; index += 1) {
            const file = pendingFiles[index];
            let lastError = null;
            for (let attempt = 1; attempt <= 2; attempt += 1) {
                try {
                    await uploadSingleFile(currentJobId, file, index, pendingFiles.length);
                    lastError = null;
                    break;
                } catch (error) {
                    lastError = error;
                    if (!error.connectionLost || attempt === 2) break;
                    await new Promise(resolve => setTimeout(resolve, 750));
                }
            }
            if (lastError) throw lastError;
        }

        await fetchJson(`/upload/finish/${currentJobId}`, { method: 'POST' });
        btn.innerHTML = '⏳ Convirtiendo...';
        document.getElementById('progressBar').style.width = '0%';
        document.getElementById('progressPercent').textContent = '0%';
        pollInterval = setInterval(pollStatus, 2500);
        await pollStatus();

    } catch (err) {
        conversionInProgress = false;
        if (currentJobId) {
            try {
                await fetch(`/upload/cancel/${currentJobId}`, { method: 'POST' });
            } catch (_cancelError) { /* The server may already be unavailable. */ }
        }
        const message = await explainConnectionFailure(err);
        document.getElementById('progressDetails').innerHTML =
            `<strong style="color:var(--red);">No se pudo completar la carga.</strong><br>${escapeHtml(message)}`;
        alert(message);
        currentJobId = null;
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
        const job = await fetchJson(`/status/${currentJobId}`, { cache: 'no-store' });
        consecutivePollFailures = 0;
        pollFailureReported = false;
        updateUI(job);
        if (job.status === 'done') {
            clearInterval(pollInterval);
            conversionInProgress = false;
            document.getElementById('convertBtn').innerHTML = '✓ Completado';
            document.getElementById('resultsActions').classList.add('active');
        }
    } catch (error) {
        consecutivePollFailures += 1;
        if (consecutivePollFailures < 3 || pollFailureReported) return;
        pollFailureReported = true;
        const message = await explainConnectionFailure(error);
        document.getElementById('progressDetails').innerHTML =
            `<strong style="color:var(--red);">Se perdió el seguimiento de la conversión.</strong><br>${escapeHtml(message)}`;
        alert(message);
    }
}

function updateUI(job) {
    const progress = job.progress || {};
    const pct = Number.isFinite(progress.percent)
        ? progress.percent
        : (job.total > 0 ? Math.round((job.completed / job.total) * 100) : 0);
    document.getElementById('progressBar').style.width = pct + '%';
    const done = Number.isFinite(progress.done) ? progress.done : job.files.filter(f => f.status === 'done').length;
    const errors = Number.isFinite(progress.errors) ? progress.errors : job.files.filter(f => f.status === 'error').length;
    const pending = Number.isFinite(progress.pending) ? progress.pending : job.files.filter(f => f.status === 'queued').length;
    document.getElementById('progressText').textContent =
        `${done} completados · ${pending} pendientes · ${errors} errores`;
    document.getElementById('progressPercent').textContent = pct + '%';

    const activeFiles = progress.active || job.files.filter(f => f.status === 'converting');
    let detailLines = [];
    activeFiles.forEach(file => {
        const page = file.current_page || 0;
        const totalPages = file.total_pages || 0;
        const pageText = totalPages ? ` · página ${page} de ${totalPages}` : '';
        const modeText = file.ocr_active || file.extraction_mode === 'ocr'
            ? ' · aplicando OCR'
            : (file.extraction_mode === 'digital_fast'
                ? ' · texto digital rápido'
                : (file.extraction_mode === 'layout' ? ' · analizando diseño' : ''));
        detailLines.push(`<strong>${escapeHtml(file.original_name)}</strong>${pageText}` +
            (modeText ? `<span class="ocr-notice">${modeText}</span>` : ''));
    });
    if (job.status === 'packaging') {
        detailLines = ['<strong>Preparando las descargas…</strong>'];
    } else if (job.status === 'done') {
        detailLines = [`<strong>Conversión terminada.</strong> ${done} archivo(s) listo(s)` +
            (errors ? ` y ${errors} con error.` : '.') +
            ' Descarga los resultados o pulsa “Convertir más documentos” para iniciar otro lote.'];
    } else if (!detailLines.length) {
        detailLines = ['Preparando la cola de conversión…'];
    }
    const elapsedText = formatDuration(progress.elapsed_seconds || 0);
    const etaText = progress.eta_seconds == null ? 'calculando…' : formatDuration(progress.eta_seconds);
    detailLines.push(`Tiempo transcurrido: ${elapsedText}` +
        (job.status === 'done' ? '' : ` · Tiempo restante aproximado: ${etaText}`));
    document.getElementById('progressDetails').innerHTML = detailLines.join('<br>');

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
        if (f.status === 'converting' && f.total_pages) {
            statusLabel += ` · pág. ${f.current_page || 0}/${f.total_pages}`;
            if (f.ocr_active || f.extraction_mode === 'ocr') statusLabel += ' · OCR';
            else if (f.extraction_mode === 'digital_fast') statusLabel += ' · digital rápido';
            else if (f.extraction_mode === 'layout') statusLabel += ' · diseño';
        }

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
        const retryButton = document.getElementById('retryErrorsBtn');
        retryButton.style.display = errors > 0 && job.status === 'done' ? 'inline-flex' : 'none';
        const totalTime = job.files.reduce((s, f) => s + (f.time || 0), 0);
        document.getElementById('statDone').textContent = done;
        document.getElementById('statErrors').textContent = errors;
        document.getElementById('statTime').textContent = totalTime.toFixed(1) + 's';
    }
}

function formatDuration(totalSeconds) {
    const seconds = Math.max(0, Math.round(Number(totalSeconds) || 0));
    if (seconds < 60) return seconds + 's';
    const minutes = Math.floor(seconds / 60);
    const remainder = seconds % 60;
    if (minutes < 60) return `${minutes} min ${remainder}s`;
    const hours = Math.floor(minutes / 60);
    return `${hours} h ${minutes % 60} min`;
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
        conversionInProgress = true;
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
    from waitress import serve

    initialize_application()
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 5000
    print(f"""
+----------------------------------------------------+
|   PDF -> Markdown  ·  Batch Converter             |
|   http://localhost:{port}                          |
|   Ctrl+C para detener                            |
+----------------------------------------------------+
    """)
    serve(app, host="127.0.0.1", port=port, threads=6)








