"""Create the illustrated PDF2MD installation and usage guide."""

from __future__ import annotations

from pathlib import Path
import sys

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    Image,
    KeepTogether,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from runtime_config import APP_VERSION  # noqa: E402


OUTPUT = ROOT / "output" / "pdf" / "Guia_rapida_PDF2MD.pdf"
ASSETS = ROOT / "assets"
PAGE_WIDTH, PAGE_HEIGHT = A4
ORANGE = colors.HexColor("#FF6B35")
DARK = colors.HexColor("#12121A")
INK = colors.HexColor("#222331")
MUTED = colors.HexColor("#68697A")
PALE = colors.HexColor("#FFF1EB")
GREEN = colors.HexColor("#168A62")


def register_fonts() -> tuple[str, str, str]:
    mono = ASSETS / "fonts" / "JetBrainsMono-Regular.ttf"
    mono_bold = ASSETS / "fonts" / "JetBrainsMono-Bold.ttf"
    if mono.is_file():
        pdfmetrics.registerFont(TTFont("JetBrainsMono", str(mono)))
    if mono_bold.is_file():
        pdfmetrics.registerFont(TTFont("JetBrainsMono-Bold", str(mono_bold)))
    return (
        "Helvetica",
        "JetBrainsMono" if mono.is_file() else "Courier",
        "JetBrainsMono-Bold" if mono_bold.is_file() else "Courier-Bold",
    )


BODY_FONT, MONO_FONT, MONO_BOLD = register_fonts()


def footer(canvas, document) -> None:
    canvas.saveState()
    canvas.setStrokeColor(colors.HexColor("#E2E2E9"))
    canvas.line(18 * mm, 14 * mm, PAGE_WIDTH - 18 * mm, 14 * mm)
    canvas.setFont(BODY_FONT, 8)
    canvas.setFillColor(MUTED)
    canvas.drawString(18 * mm, 9.5 * mm, f"PDF2MD {APP_VERSION} - Procesamiento completamente local")
    canvas.drawRightString(PAGE_WIDTH - 18 * mm, 9.5 * mm, f"Página {document.page}")
    canvas.restoreState()


def step_card(number: str, title: str, body: str, hint: str = ""):
    number_style = ParagraphStyle(
        "number", fontName=MONO_BOLD, fontSize=18, textColor=colors.white,
        alignment=TA_CENTER, leading=22,
    )
    title_style = ParagraphStyle(
        "step-title", fontName=BODY_FONT, fontSize=13, leading=16,
        textColor=INK, spaceAfter=4,
    )
    body_style = ParagraphStyle(
        "step-body", fontName=BODY_FONT, fontSize=9.5, leading=13.5,
        textColor=MUTED,
    )
    rows = [[
        Paragraph(number, number_style),
        [
            Paragraph(f"<b>{title}</b>", title_style),
            Paragraph(body, body_style),
            Paragraph(f"<font color='#168A62'>{hint}</font>", body_style) if hint else Spacer(1, 0),
        ],
    ]]
    table = Table(rows, colWidths=[18 * mm, 142 * mm], hAlign="LEFT")
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), ORANGE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (0, 0), 5),
        ("RIGHTPADDING", (0, 0), (0, 0), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING", (1, 0), (1, 0), 12),
        ("RIGHTPADDING", (1, 0), (1, 0), 12),
        ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#DFDFE8")),
    ]))
    return table


def build() -> Path:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "title", parent=styles["Title"], fontName=BODY_FONT,
        fontSize=25, leading=29, textColor=DARK, alignment=TA_LEFT,
        spaceAfter=4,
    )
    subtitle = ParagraphStyle(
        "subtitle", parent=styles["BodyText"], fontName=BODY_FONT,
        fontSize=11, leading=15, textColor=MUTED, spaceAfter=14,
    )
    section = ParagraphStyle(
        "section", parent=styles["Heading2"], fontName=BODY_FONT,
        fontSize=17, leading=21, textColor=DARK, spaceBefore=4, spaceAfter=10,
    )
    body = ParagraphStyle(
        "body", parent=styles["BodyText"], fontName=BODY_FONT,
        fontSize=10, leading=14, textColor=INK, spaceAfter=8,
    )
    callout = ParagraphStyle(
        "callout", parent=body, backColor=PALE, borderColor=ORANGE,
        borderWidth=0.8, borderPadding=10, textColor=INK, spaceBefore=7, spaceAfter=9,
    )

    frame = Frame(18 * mm, 18 * mm, PAGE_WIDTH - 36 * mm, PAGE_HEIGHT - 34 * mm, id="content")
    document = BaseDocTemplate(
        str(OUTPUT), pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=18 * mm, bottomMargin=18 * mm,
        title="Guía rápida de PDF2MD", author="PDF2MD",
    )
    document.addPageTemplates(PageTemplate(id="guide", frames=[frame], onPage=footer))

    logo = ASSETS / "pdf2md.png"
    story = []
    if logo.is_file():
        story.append(Image(str(logo), width=22 * mm, height=22 * mm, hAlign="LEFT"))
        story.append(Spacer(1, 3 * mm))
    story.extend([
        Paragraph("Guía rápida de PDF2MD", title),
        Paragraph("Instala una vez. Después convierte documentos desde tu navegador con un clic.", subtitle),
        Paragraph("Instalación", section),
        step_card("1", "Descarga el instalador", f"Abre el enlace enviado por tu docente y descarga <b>Instalar_PDF2MD_{APP_VERSION}.exe</b>."),
        Spacer(1, 4 * mm),
        step_card("2", "Ejecuta el archivo", "Haz doble clic en el instalador y sigue los botones de la ventana. No necesitas instalar Python ni Tesseract.", "La instalación funciona sin permisos de administrador."),
        Spacer(1, 4 * mm),
        step_card("3", "Abre PDF2MD", "Al finalizar, marca <b>Abrir PDF2MD</b>. También tendrás un acceso directo en el escritorio."),
        Spacer(1, 5 * mm),
        Paragraph(
            "<b>Posible aviso de Windows:</b> esta versión piloto no está firmada digitalmente. "
            "Si aparece Windows protegió su PC, confirma primero que el nombre del archivo y el enlace "
            "coinciden con los enviados por tu docente. Después selecciona <b>Más información</b> y "
            "<b>Ejecutar de todas formas</b>.", callout,
        ),
        Paragraph("PDF2MD se abre en el navegador, pero funciona dentro de tu computador. No necesitas conexión para convertir documentos.", body),
        PageBreak(),
        Paragraph("Convertir documentos", title),
        Paragraph("La interfaz se mantiene igual: selecciona, convierte y descarga.", subtitle),
        step_card("1", "Selecciona archivos", "Arrastra tus PDF, Word o RTF al área principal, o haz clic para buscarlos. Puedes agregar varios a la vez."),
        Spacer(1, 4 * mm),
        step_card("2", "Convierte", "Haz clic en <b>Convertir todo</b>. El progreso muestra archivos pendientes, página actual, OCR y tiempo aproximado."),
        Spacer(1, 4 * mm),
        step_card("3", "Descarga", "Al terminar puedes descargar cada Markdown, un único documento unido o todos los resultados en ZIP."),
        Spacer(1, 4 * mm),
        step_card("4", "Convierte más documentos", "Pulsa <b>Convertir más documentos</b> y selecciona el siguiente lote. No necesitas actualizar la página del navegador."),
        Spacer(1, 5 * mm),
        Paragraph("¿Cuándo usar OCR?", section),
        Paragraph(
            "<b>Auto-detectar</b> es la opción recomendada. Los PDF con texto seleccionable usan extracción rápida; "
            "PDF2MD activa OCR solo en páginas escaneadas. "
            "Usa <b>Forzar OCR</b> cuando el documento sea una imagen o el texto no pueda seleccionarse. "
            "Usa <b>Desactivado</b> únicamente para PDF digitales.", body,
        ),
        Paragraph(
            "<b>Privacidad:</b> los documentos se procesan localmente y no se envían a internet. "
            "Las descargas quedan en la carpeta configurada por tu navegador.", callout,
        ),
        Paragraph(
            "Para ayuda adicional, abre <b>Acerca de y ayuda</b> dentro de PDF2MD. Allí encontrarás "
            "soluciones frecuentes, la versión instalada y el acceso a la carpeta de diagnóstico.", body,
        ),
    ])
    document.build(story)
    return OUTPUT


if __name__ == "__main__":
    print(build())
