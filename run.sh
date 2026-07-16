#!/bin/bash
# ─── PDF2MD Batch Converter - Setup & Run ───
echo "📄 PDF → Markdown Batch Converter"
echo "=================================="

# Check Python
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 no encontrado. Instálalo primero."
    exit 1
fi

# Install dependencies
echo "📦 Instalando dependencias..."
pip install pymupdf4llm flask 2>/dev/null || pip install pymupdf4llm flask --break-system-packages 2>/dev/null

if [ $? -ne 0 ]; then
    echo "❌ Error instalando dependencias. Intenta:"
    echo "   pip install pymupdf4llm flask"
    exit 1
fi

echo "✅ Dependencias instaladas"
echo ""

# Run
python3 app.py "$@"
