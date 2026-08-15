#!/usr/bin/env python3
"""Mirror a PDF directory tree as Markdown using the app's converter.

This script intentionally delegates PDF extraction and OCR decisions to
``app.pdf_to_markdown_page_by_page``.  It only handles discovery, mirrored
paths, resumability, atomic writes, progress, and error reporting.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app import TESSDATA_DIR, pdf_to_markdown_page_by_page  # noqa: E402


def atomic_write_markdown(pdf_path: Path, md_path: Path, options: dict) -> int:
    md_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = md_path.with_suffix(md_path.suffix + ".partial")
    markdown = pdf_to_markdown_page_by_page(pdf_path, md_path, options)
    if not markdown.strip():
        raise RuntimeError("La conversión produjo Markdown vacío")
    temporary_path.write_text(markdown, encoding="utf-8")
    os.replace(temporary_path, md_path)
    return len(markdown)


def write_progress(path: Path, data: dict) -> None:
    temporary_path = path.with_suffix(path.suffix + ".partial")
    payload = json.dumps(data, ensure_ascii=False, indent=2)
    last_error = None
    for attempt in range(10):
        try:
            temporary_path.write_text(payload, encoding="utf-8")
            os.replace(temporary_path, path)
            return
        except PermissionError as exc:
            last_error = exc
            time.sleep(0.25 * (attempt + 1))
    # OneDrive can briefly lock the destination while syncing. Progress is
    # informational, so fall back to a direct write after bounded retries.
    try:
        path.write_text(payload, encoding="utf-8")
    except PermissionError:
        print(f"AVISO: no se pudo actualizar progreso: {last_error}", flush=True)


def append_result(path: Path, row: list[str]) -> None:
    new_file = not path.exists()
    with path.open("a", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        if new_file:
            writer.writerow(
                ["fecha", "estado", "pdf_relativo", "md_relativo", "segundos", "detalle"]
            )
        writer.writerow(row)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--ocr-mode", choices=("auto", "force", "none"), default="auto")
    parser.add_argument("--ocr-dpi", type=int, default=300)
    parser.add_argument("--extract-images", action="store_true")
    parser.add_argument("--allow-empty-ocr-pages", action="store_true")
    parser.add_argument(
        "--skip-relative",
        action="append",
        default=[],
        help="Ruta PDF relativa que se registrará y omitirá temporalmente",
    )
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    source = args.source.expanduser().resolve()
    output = args.output.expanduser().resolve()
    if not source.is_dir():
        raise SystemExit(f"La carpeta de origen no existe: {source}")
    if output == source or source not in output.parents:
        raise SystemExit("La salida debe ser una subcarpeta del origen")

    output.mkdir(parents=True, exist_ok=True)
    progress_path = output / "_conversion_progress.json"
    results_path = output / "_conversion_results.csv"

    pdfs = sorted(
        (
            path
            for path in source.rglob("*")
            if path.is_file()
            and path.suffix.lower() == ".pdf"
            and output not in path.parents
        ),
        key=lambda path: str(path.relative_to(source)).lower(),
    )
    total = len(pdfs)
    options = {
        "ocr_mode": args.ocr_mode,
        "ocr_dpi": max(72, min(args.ocr_dpi, 600)),
        "extract_images": args.extract_images,
        "exclude_headers": False,
        "allow_empty_ocr_pages": args.allow_empty_ocr_pages,
    }

    started = time.time()
    done = skipped = failed = 0
    explicit_skips = {str(Path(value)).lower() for value in args.skip_relative}
    print(f"PDF encontrados: {total}", flush=True)
    print(f"Salida: {output}", flush=True)
    print(f"Tessdata: {TESSDATA_DIR or 'NO DETECTADO'}", flush=True)

    for index, pdf_path in enumerate(pdfs, start=1):
        relative_pdf = pdf_path.relative_to(source)
        relative_md = relative_pdf.with_suffix(".md")
        md_path = output / relative_md
        item_started = time.time()
        state = ""
        detail = ""

        try:
            if str(relative_pdf).lower() in explicit_skips:
                skipped += 1
                state = "omitido_explicito"
                detail = "Excluido temporalmente para recuperación independiente"
            elif md_path.exists() and md_path.stat().st_size > 0 and not args.overwrite:
                skipped += 1
                state = "omitido_existente"
                detail = "Markdown no vacío ya existente"
            else:
                char_count = atomic_write_markdown(pdf_path, md_path, options)
                done += 1
                state = "convertido"
                detail = f"{char_count} caracteres"
        except Exception as exc:
            failed += 1
            state = "error"
            detail = f"{type(exc).__name__}: {exc}"
            partial_path = md_path.with_suffix(md_path.suffix + ".partial")
            if partial_path.exists():
                partial_path.unlink()
            with (output / "_conversion_errors.log").open("a", encoding="utf-8") as stream:
                stream.write(f"\n[{datetime.now().isoformat()}] {relative_pdf}\n")
                stream.write(traceback.format_exc())

        elapsed = time.time() - item_started
        append_result(
            results_path,
            [
                datetime.now().isoformat(timespec="seconds"),
                state,
                str(relative_pdf),
                str(relative_md),
                f"{elapsed:.2f}",
                detail,
            ],
        )
        progress = {
            "source": str(source),
            "output": str(output),
            "total": total,
            "processed": index,
            "converted": done,
            "skipped": skipped,
            "failed": failed,
            "current": str(relative_pdf),
            "ocr_mode": args.ocr_mode,
            "ocr_dpi": options["ocr_dpi"],
            "elapsed_seconds": round(time.time() - started, 1),
            "updated_at": datetime.now().isoformat(timespec="seconds"),
            "finished": index == total,
        }
        write_progress(progress_path, progress)
        print(
            f"[{index}/{total}] {state}: {relative_pdf} ({elapsed:.1f}s)",
            flush=True,
        )

    print(
        f"FINAL: convertidos={done}, omitidos={skipped}, errores={failed}, "
        f"segundos={time.time() - started:.1f}",
        flush=True,
    )
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
