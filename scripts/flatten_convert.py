#!/usr/bin/env python3
"""Convertir recursivamente PDFs a Markdown y guardar en carpeta plana.

Usage:
    python scripts/flatten_convert.py --input "C:\\path\\to\\dir" --output "C:\\path\\to\\out" [--ocr auto|force|none]

The script uses `pymupdf4llm.to_markdown` (already used by the project).
"""
import argparse
import sys
import time
from pathlib import Path
import shutil

import fitz
import pymupdf4llm


def is_pdf_scanned(pdf_path: Path) -> bool:
    try:
        doc = fitz.open(str(pdf_path))
        for i in range(min(5, len(doc))):
            page = doc[i]
            if hasattr(page, "is_searchable") and page.is_searchable():
                return False
            text = page.get_text().strip()
            if len(text) > 100:
                return False
        return True
    except Exception:
        return True
    finally:
        if 'doc' in locals():
            doc.close()


def unique_name(dest_dir: Path, base_name: str):
    candidate = dest_dir / base_name
    if not candidate.exists():
        return candidate
    stem = Path(base_name).stem
    suffix = Path(base_name).suffix
    i = 1
    while True:
        new_name = f"{stem}_{i}{suffix}"
        candidate = dest_dir / new_name
        if not candidate.exists():
            return candidate
        i += 1


def convert_file(pdf_path: Path, out_dir: Path, ocr_mode: str = "auto"):
    out_dir.mkdir(parents=True, exist_ok=True)
    md_name = pdf_path.stem + ".md"
    md_path = unique_name(out_dir, md_name)

    use_ocr = False
    if ocr_mode == "force":
        use_ocr = True
    elif ocr_mode == "auto":
        use_ocr = is_pdf_scanned(pdf_path)

    images_dir = out_dir / "images"
    images_dir.mkdir(parents=True, exist_ok=True)

    try:
        md_text = pymupdf4llm.to_markdown(
            str(pdf_path),
            write_images=True,
            image_path=str(images_dir),
            image_format="png",
            dpi=150,
            force_ocr=use_ocr,
        )
        md_path.write_text(md_text, encoding="utf-8")
        return True, md_path
    except Exception as e:
        return False, str(e)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="Carpeta raíz con PDFs (se recorre recursivamente)")
    parser.add_argument("--output", required=True, help="Carpeta de salida donde se guardarán todos los .md (plana)")
    parser.add_argument("--ocr", choices=["auto", "force", "none"], default="auto", help="Modo OCR")
    args = parser.parse_args()

    input_dir = Path(args.input).expanduser().resolve()
    output_dir = Path(args.output).expanduser().resolve()

    if not input_dir.exists() or not input_dir.is_dir():
        print(f"Error: input dir no existe: {input_dir}")
        sys.exit(2)

    output_dir.mkdir(parents=True, exist_ok=True)

    pdfs = list(input_dir.rglob("*.pdf"))
    if not pdfs:
        print("No se encontraron archivos PDF en la ruta especificada.")
        return

    print(f"Encontrados {len(pdfs)} PDFs. Convirtiendo a: {output_dir}")

    succeeded = 0
    failed = 0
    start = time.time()

    for p in pdfs:
        rel = p.relative_to(input_dir)
        print(f"-> {rel}")
        ok, result = convert_file(p, output_dir, args.ocr)
        if ok:
            print(f"   ✔ generado: {result.name}")
            succeeded += 1
        else:
            print(f"   ✖ error: {result}")
            failed += 1

    total_time = time.time() - start
    print("\nResumen:")
    print(f"  Convertidos: {succeeded}")
    print(f"  Errores: {failed}")
    print(f"  Tiempo total: {total_time:.1f}s")


if __name__ == "__main__":
    main()
