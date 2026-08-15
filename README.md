# PDF2MD · Conversor documental

Aplicación web local para transformar documentos entre PDF, Word, RTF y Markdown.

## Funciones principales

- Convierte PDF, DOCX, DOCM y RTF a Markdown estándar.
- Conserva títulos, negrita, cursiva, listas, tablas, citas, callouts, enlaces y bloques de código.
- Permite descargar archivos individuales, documentos combinados o paquetes ZIP.
- Ejecuta todo localmente: los documentos no se envían a servicios externos.

## Instalación para estudiantes en Windows

1. Descarga `Instalar_PDF2MD_1.0.3.exe`.
2. Abre el instalador y sigue sus indicaciones.
3. Al terminar, PDF2MD se abrirá automáticamente en el navegador.
4. En los usos siguientes, abre el acceso directo **PDF2MD** del escritorio.

El instalador es autocontenido: incluye Python, OCR en español/inglés y todas
las librerías. No utiliza `pip`, `winget` ni descargas durante la instalación.
Los documentos se procesan exclusivamente en el computador del usuario.

## Ejecución manual

~~~bash
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python app.py
~~~

La ejecución manual es solo para desarrollo. La distribución estudiantil se
construye mediante:

~~~powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\build_release.ps1
~~~

El resultado queda en `release/` junto con la guía, versión, licencias y
checksum SHA-256.

En Linux o macOS:

~~~bash
chmod +x run.sh
./run.sh
~~~

## Uso

### PDF, Word o RTF → Markdown

1. Abre la pestaña **Convertir archivos**.
2. Agrega PDF, DOCX, DOCM o RTF.
3. Ajusta OCR e imágenes si lo necesitas.
4. Convierte y descarga los MD estándar individuales, el documento combinado o el ZIP.

### Markdown → PDF o Word

1. Abre la pestaña **MD → PDF / Word**.
2. Agrega uno o varios archivos .md o .markdown.
3. Si las notas usan imágenes locales, agrégalas en la misma selección.
4. Elige PDF, Word o ambos.
5. Descarga cada archivo o el ZIP completo.

### Unir Markdown

La pestaña **Unir MDs** combina varios archivos en un único documento Markdown.

## Puerto personalizado

~~~bash
iniciar_app.bat 8080
~~~

o:

~~~bash
python app.py 8080
~~~

## Tecnologías

- PyMuPDF4LLM y PyMuPDF para PDF y OCR.
- python-docx para lectura y escritura de Word.
- striprtf para extraer texto de documentos RTF.
- Python-Markdown y Beautiful Soup para interpretar y maquetar Markdown.
- Flask para la interfaz web local.
