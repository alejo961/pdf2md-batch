# PDF2MD · Conversor documental

Aplicación web local para transformar documentos entre PDF, Word y Markdown, con soporte específico para Obsidian.

## Funciones principales

- Convierte PDF, DOCX y DOCM a Markdown estándar.
- Genera simultáneamente una nota Obsidian Flavored Markdown (.obsidian.md).
- Las notas de Obsidian incluyen propiedades YAML, tags, alias, wikilinks y embeds compatibles.
- Convierte Markdown estándar u Obsidian a PDF maquetado y buscable.
- Convierte Markdown estándar u Obsidian a Word editable (.docx).
- Conserva títulos, negrita, cursiva, listas, tablas, citas, callouts, enlaces y bloques de código.
- Permite descargar archivos individuales, documentos combinados o paquetes ZIP.
- Ejecuta todo localmente: los documentos no se envían a servicios externos.

## Instalación rápida en Windows

1. Descarga el repositorio mediante **Code → Download ZIP**.
2. Descomprime la carpeta completa.
3. Ejecuta iniciar_app.bat.
4. Espera la instalación inicial de Python y las dependencias.
5. Abre http://localhost:5000 si el navegador no se inicia automáticamente.

No cierres la ventana de comandos mientras uses la aplicación. Para detenerla, presiona Ctrl+C.

## Ejecución manual

~~~bash
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python app.py
~~~

En Linux o macOS:

~~~bash
chmod +x run.sh
./run.sh
~~~

## Uso

### PDF o Word → Markdown

1. Abre la pestaña **Convertir archivos**.
2. Agrega PDF, DOCX o DOCM.
3. Ajusta OCR e imágenes si lo necesitas.
4. Convierte y elige entre:
   - **MD estándar**, compatible con CommonMark/GitHub.
   - **Obsidian MD**, con frontmatter, wikilinks y embeds.
   - ZIP con ambas variantes y sus imágenes.

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
- Python-Markdown y Beautiful Soup para interpretar y maquetar Markdown.
- Flask para la interfaz web local.
