from __future__ import annotations

import io
import os
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz
from PIL import Image, ImageDraw, ImageFont


TEST_TEMP_ROOT = Path(__file__).resolve().parents[1] / ".tmp" / "tests" / "manual"
TEST_TEMP_ROOT.mkdir(parents=True, exist_ok=True)
tempfile.tempdir = str(TEST_TEMP_ROOT)
TEST_DATA = TEST_TEMP_ROOT / "runtime"
TEST_DATA.mkdir(parents=True, exist_ok=True)
os.environ["PDF2MD_DATA_DIR"] = str(TEST_DATA)

import app as pdf2md  # noqa: E402


class PDF2MDApplicationTests(unittest.TestCase):
    @classmethod
    def tearDownClass(cls):
        pdf2md.conversion_executor.shutdown(wait=True, cancel_futures=True)

    def test_health_help_and_local_assets(self):
        client = pdf2md.app.test_client()
        health = client.get("/health")
        self.assertEqual(health.status_code, 200)
        self.assertEqual(health.get_json()["status"], "ok")
        page = client.get("/")
        self.assertEqual(page.status_code, 200)
        html = page.get_data(as_text=True)
        self.assertIn("Acerca de y ayuda", html)
        self.assertIn("progressDetails", html)
        self.assertIn("Copiar información de diagnóstico", html)
        self.assertIn("Convertir más documentos", html)
        self.assertIn("function chooseMoreDocuments()", html)
        self.assertIn("function resetConversionView()", html)
        self.assertIn("sin recargar la página", html)
        self.assertNotIn("fonts.googleapis.com", html)
        font = client.get("/assets/fonts/Outfit-Variable.ttf")
        self.assertEqual(font.status_code, 200)
        font.close()

        diagnostic = client.get("/diagnostics")
        self.assertEqual(diagnostic.status_code, 200)
        self.assertIn(f"PDF2MD {pdf2md.APP_VERSION}", diagnostic.get_data(as_text=True))

    def test_incremental_upload_avoids_large_batch_request(self):
        client = pdf2md.app.test_client()
        previous_limit = pdf2md.app.config["MAX_CONTENT_LENGTH"]
        pdf2md.app.config["MAX_CONTENT_LENGTH"] = 1200
        payloads = [(io.BytesIO(b"%PDF-1.4\n" + b"x" * 256), f"parte-{index}.pdf") for index in range(5)]
        try:
            oversized = client.post(
                "/upload",
                data={"files": payloads},
                content_type="multipart/form-data",
            )
            self.assertEqual(oversized.status_code, 413)
            self.assertIn("límite", oversized.get_json()["error"])

            started = client.post("/upload/start", json={"ocrMode": "none"})
            self.assertEqual(started.status_code, 200)
            job_id = started.get_json()["job_id"]
            for index in range(5):
                uploaded = client.post(
                    f"/upload/file/{job_id}",
                    data={
                        "file": (
                            io.BytesIO(b"%PDF-1.4\n" + b"x" * 256),
                            f"parte-{index}.pdf",
                        )
                    },
                    content_type="multipart/form-data",
                )
                self.assertEqual(uploaded.status_code, 200)
            self.assertEqual(pdf2md.jobs[job_id]["total"], 5)
            cancelled = client.post(f"/upload/cancel/{job_id}")
            self.assertEqual(cancelled.get_json()["status"], "cancelled")
        finally:
            pdf2md.app.config["MAX_CONTENT_LENGTH"] = previous_limit

    def test_digital_pdf_conversion_and_progress(self):
        source = TEST_TEMP_ROOT / "digital.pdf"
        output = TEST_TEMP_ROOT / "digital.md"
        if source.exists():
            source.unlink()
        document = fitz.open()
        page = document.new_page()
        page.insert_text(
            (72, 90),
            "Normativa aduanera de Colombia con texto digital completamente seleccionable",
            fontsize=14,
        )
        document.save(source)
        document.close()

        updates = []
        with patch.object(
            pdf2md.pymupdf4llm,
            "to_markdown",
            side_effect=AssertionError("Un PDF digital simple no debe activar el motor de diseño"),
        ):
            markdown = pdf2md.pdf_to_markdown_page_by_page(
                source,
                output,
                {"ocr_mode": "auto", "ocr_dpi": 300},
                progress_callback=lambda **values: updates.append(values),
            )
        self.assertIn("Normativa aduanera", markdown)
        self.assertEqual(updates[-1]["current_page"], 1)
        self.assertFalse(updates[-1]["ocr_active"])
        self.assertEqual(updates[-1]["extraction_mode"], "digital_fast")

    def test_rtf_conversion_preserves_accents_paragraphs_and_bullets(self):
        rtf = (
            rb"{\rtf1\ansi\ansicpg1252 "
            rb"Reglamentaci\'f3n aduanera\par "
            rb"\bullet\tab Primer requisito\par "
            rb"Segunda l\u237?nea}"
        )
        markdown = pdf2md.rtf_document_to_markdown(rtf)
        self.assertIn("Reglamentación aduanera", markdown)
        self.assertIn("- Primer requisito", markdown)
        self.assertIn("Segunda línea", markdown)
        self.assertTrue(pdf2md.is_supported_convert_file("documento.RTF"))

    @unittest.skipUnless(pdf2md.TESSDATA_DIR, "Tesseract spa+eng no está disponible")
    def test_scanned_pdf_uses_spanish_ocr(self):
        image_path = TEST_TEMP_ROOT / "scan.png"
        pdf_path = TEST_TEMP_ROOT / "scan.pdf"
        output = TEST_TEMP_ROOT / "scan.md"
        if pdf_path.exists():
            pdf_path.unlink()
        image = Image.new("RGB", (1600, 900), "white")
        draw = ImageDraw.Draw(image)
        font_path = Path(__file__).resolve().parents[1] / "assets" / "fonts" / "Outfit-Variable.ttf"
        font = ImageFont.truetype(str(font_path), 72)
        draw.text((100, 260), "NORMA LEGAL COLOMBIA", font=font, fill="black")
        image.save(image_path)

        document = fitz.open()
        page = document.new_page(width=800, height=450)
        page.insert_image(page.rect, filename=str(image_path))
        document.save(pdf_path)
        document.close()

        updates = []
        markdown = pdf2md.pdf_to_markdown_page_by_page(
            pdf_path,
            output,
            {"ocr_mode": "auto", "ocr_dpi": 300},
            progress_callback=lambda **values: updates.append(values),
        )
        self.assertIn("COLOMBIA", markdown.upper())
        self.assertTrue(any(item.get("ocr_active") for item in updates))

    @unittest.skipUnless(pdf2md.TESSDATA_DIR, "Tesseract spa+eng no está disponible")
    def test_mixed_pdf_uses_fast_text_and_page_ocr(self):
        image_path = TEST_TEMP_ROOT / "mixed-scan.png"
        pdf_path = TEST_TEMP_ROOT / "mixed.pdf"
        image = Image.new("RGB", (1400, 700), "white")
        draw = ImageDraw.Draw(image)
        font_path = Path(__file__).resolve().parents[1] / "assets" / "fonts" / "Outfit-Variable.ttf"
        font = ImageFont.truetype(str(font_path), 64)
        draw.text((90, 230), "PAGINA ESCANEADA COLOMBIA", font=font, fill="black")
        image.save(image_path)

        document = fitz.open()
        digital_page = document.new_page(width=800, height=450)
        digital_page.insert_text(
            (60, 90),
            "Primera página con suficiente contenido digital seleccionable para la ruta rápida",
            fontsize=14,
        )
        scanned_page = document.new_page(width=800, height=400)
        scanned_page.insert_image(scanned_page.rect, filename=str(image_path))
        document.save(pdf_path)
        document.close()

        updates = []
        markdown = pdf2md.pdf_to_markdown_page_by_page(
            pdf_path,
            TEST_TEMP_ROOT / "mixed.md",
            {"ocr_mode": "auto", "ocr_dpi": 300},
            progress_callback=lambda **values: updates.append(values),
        )
        modes = {item.get("extraction_mode") for item in updates}
        self.assertIn("digital_fast", modes)
        self.assertIn("ocr", modes)
        self.assertIn("COLOMBIA", markdown.upper())

    def test_upload_job_reaches_completion(self):
        document = fitz.open()
        page = document.new_page()
        page.insert_text(
            (72, 90),
            "Documento jurídico digital con contenido seleccionable para conversión rápida",
            fontsize=14,
        )
        payload = document.tobytes()
        document.close()

        client = pdf2md.app.test_client()
        response = client.post(
            "/upload",
            data={
                "files": (io.BytesIO(payload), "prueba jurídica.pdf"),
                "ocrMode": "auto",
                "ocrDpi": "300",
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 200)
        job_id = response.get_json()["job_id"]
        deadline = time.time() + 30
        status = None
        while time.time() < deadline:
            status = client.get(f"/status/{job_id}").get_json()
            if status.get("status") == "done":
                break
            time.sleep(0.2)
        self.assertEqual(status.get("status"), "done")
        self.assertEqual(status["progress"]["done"], 1)
        self.assertEqual(status["progress"]["percent"], 100)


if __name__ == "__main__":
    unittest.main(verbosity=2)
