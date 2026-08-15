"""Offline dependency and contents audit for the packaged Windows application."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
from pathlib import Path

import pefile


PE_SUFFIXES = {".exe", ".dll", ".pyd"}
SYSTEM_PREFIXES = ("api-ms-win-", "ext-ms-")
SOURCE_FILES = ("app.py", "launcher.py", "runtime_config.py", "conversion_formats.py")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def system_dll_names() -> set[str]:
    names = set()
    windows = Path(os.environ.get("WINDIR", r"C:\Windows"))
    for directory in (windows / "System32", windows / "SysWOW64"):
        if not directory.is_dir():
            continue
        try:
            names.update(item.name.casefold() for item in directory.iterdir() if item.is_file())
        except OSError:
            continue
    return names


def imported_dlls(path: Path) -> tuple[list[str], str | None]:
    try:
        image = pefile.PE(str(path), fast_load=True)
        image.parse_data_directories(
            directories=[pefile.DIRECTORY_ENTRY["IMAGE_DIRECTORY_ENTRY_IMPORT"]]
        )
        imports = [
            entry.dll.decode("ascii", errors="replace")
            for entry in getattr(image, "DIRECTORY_ENTRY_IMPORT", [])
        ]
        image.close()
        return sorted(set(imports), key=str.casefold), None
    except pefile.PEFormatError as exc:
        return [], str(exc)


def external_source_urls(project_root: Path) -> list[str]:
    pattern = re.compile(r"https?://[^\s'\"<>]+", re.IGNORECASE)
    urls = set()
    for relative_name in SOURCE_FILES:
        path = project_root / relative_name
        if not path.is_file():
            continue
        for match in pattern.findall(path.read_text(encoding="utf-8", errors="replace")):
            cleaned = match.rstrip(".,);}")
            if "127.0.0.1" not in cleaned and "localhost" not in cleaned:
                urls.add(cleaned)
    return sorted(urls)


def audit(bundle: Path, project_root: Path) -> dict:
    bundle = bundle.resolve()
    all_files = [item for item in bundle.rglob("*") if item.is_file()]
    bundle_names = {item.name.casefold() for item in all_files}
    system_names = system_dll_names()
    pe_files = [item for item in all_files if item.suffix.casefold() in PE_SUFFIXES]

    unresolved = {}
    parse_errors = {}
    import_counts = {}
    for image_path in pe_files:
        imports, error = imported_dlls(image_path)
        relative = str(image_path.relative_to(bundle))
        if error:
            parse_errors[relative] = error
            continue
        import_counts[relative] = len(imports)
        missing = []
        for name in imports:
            folded = name.casefold()
            if folded.startswith(SYSTEM_PREFIXES):
                continue
            if folded not in bundle_names and folded not in system_names:
                missing.append(name)
        if missing:
            unresolved[relative] = missing

    required_relative = [
        "PDF2MD.exe",
        "_internal/tesseract/tessdata/spa.traineddata",
        "_internal/tesseract/tessdata/eng.traineddata",
        "_internal/tesseract/tessdata/osd.traineddata",
        "_internal/assets/pdf2md.ico",
        "_internal/assets/fonts/Outfit-Variable.ttf",
    ]
    missing_required = [name for name in required_relative if not (bundle / name).is_file()]
    executable = bundle / "PDF2MD.exe"
    result = {
        "bundle": str(bundle),
        "file_count": len(all_files),
        "size_bytes": sum(item.stat().st_size for item in all_files),
        "pe_file_count": len(pe_files),
        "main_executable_sha256": sha256(executable) if executable.is_file() else None,
        "missing_required_files": missing_required,
        "unresolved_imports": unresolved,
        "pe_parse_errors": parse_errors,
        "external_source_urls": external_source_urls(project_root),
        "import_counts": import_counts,
    }
    result["passed"] = not (
        result["missing_required_files"]
        or result["unresolved_imports"]
        or result["pe_parse_errors"]
        or result["external_source_urls"]
    )
    return result


def text_report(result: dict) -> str:
    lines = [
        "AUDITORIA TECNICA OFFLINE - PDF2MD",
        "=" * 40,
        f"Paquete: {result['bundle']}",
        f"Archivos: {result['file_count']}",
        f"Tamaño: {result['size_bytes']} bytes",
        f"Binarios PE revisados: {result['pe_file_count']}",
        f"SHA-256 del ejecutable interno: {result['main_executable_sha256']}",
        f"Resultado: {'APROBADO' if result['passed'] else 'REQUIERE REVISION'}",
        "",
        "Archivos esenciales faltantes:",
        json.dumps(result["missing_required_files"], ensure_ascii=False, indent=2),
        "",
        "Dependencias DLL no resueltas:",
        json.dumps(result["unresolved_imports"], ensure_ascii=False, indent=2),
        "",
        "Errores al analizar binarios:",
        json.dumps(result["pe_parse_errors"], ensure_ascii=False, indent=2),
        "",
        "Direcciones externas encontradas en el código de ejecución:",
        json.dumps(result["external_source_urls"], ensure_ascii=False, indent=2),
    ]
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--bundle", type=Path, default=Path("dist/PDF2MD"))
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    project_root = Path(__file__).resolve().parents[1]
    result = audit(args.bundle, project_root)
    report = text_report(result)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(report, encoding="utf-8")
    print(report, end="")
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    sys.exit(main())
