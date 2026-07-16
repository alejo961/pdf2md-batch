@echo off
setlocal
cd /d "%~dp0"

set "PORT=%~1"
if "%PORT%"=="" set "PORT=5000"

echo.
echo PDF2MD Batch Converter
echo Carpeta actual: %CD%
echo.

if not exist "app.py" (
    echo No se encontro app.py en esta carpeta.
    echo Asegurate de ejecutar este archivo desde la carpeta completa del proyecto.
    pause
    exit /b 1
)

if not exist "requirements.txt" (
    echo No se encontro requirements.txt en esta carpeta.
    echo Asegurate de compartir la carpeta completa del proyecto.
    pause
    exit /b 1
)

if not exist ".tmp" mkdir ".tmp"
set "TEMP=%CD%\.tmp"
set "TMP=%CD%\.tmp"

call :detect_python

if "%PYTHON_CMD%"=="" (
    echo Python no encontrado. Intentando instalar con winget...
    where winget >nul 2>nul
    if errorlevel 1 (
        echo winget no esta disponible en este equipo.
        start https://www.python.org/downloads/
        echo Instala Python 3 y vuelve a ejecutar este archivo.
        pause
        exit /b 1
    )

    winget install -e --id Python.Python.3.12
    call :detect_python
)

if "%PYTHON_CMD%"=="" (
    echo No fue posible detectar Python despues de la instalacion.
    start https://www.python.org/downloads/
    pause
    exit /b 1
)

for /f "usebackq delims=" %%P in (`%PYTHON_CMD% -c "import sys; print(sys.executable)"`) do set "PYTHON_EXE=%%P"
if "%PYTHON_EXE%"=="" (
    echo No fue posible resolver la ruta de Python.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Creando entorno virtual local...
    "%PYTHON_EXE%" -m venv .venv
    if errorlevel 1 (
        echo No se pudo crear el entorno virtual. Se usara Python del usuario.
        set "USE_USER_PYTHON=1"
    )
)

set "APP_PY=.venv\Scripts\python.exe"
if defined USE_USER_PYTHON set "APP_PY=%PYTHON_EXE%"

if not exist "%APP_PY%" (
    echo El entorno virtual quedo incompleto. Limpiando y reintentando...
    if exist ".venv" rmdir /s /q ".venv"
    "%PYTHON_EXE%" -m venv .venv
    if errorlevel 1 (
        echo No se pudo crear el entorno virtual. Se usara Python del usuario.
        set "APP_PY=%PYTHON_EXE%"
    ) else (
        set "APP_PY=.venv\Scripts\python.exe"
    )
)

if not exist "%APP_PY%" (
    echo No se pudo preparar Python para ejecutar la app.
    pause
    exit /b 1
)

"%APP_PY%" -c "import flask, pymupdf4llm, fitz" >nul 2>nul
if not errorlevel 1 goto deps_ready

"%APP_PY%" -m pip --version >nul 2>nul
if errorlevel 1 (
    if not "%APP_PY%"=="%PYTHON_EXE%" (
        echo El entorno virtual no tiene pip. Limpiando y reintentando...
        if exist ".venv" rmdir /s /q ".venv"
        "%PYTHON_EXE%" -m venv .venv
        if not errorlevel 1 set "APP_PY=.venv\Scripts\python.exe"
    )
)

"%APP_PY%" -c "import flask, pymupdf4llm, fitz" >nul 2>nul
if not errorlevel 1 goto deps_ready

echo Instalando/validando dependencias...
"%APP_PY%" -m pip install --upgrade pip
if errorlevel 1 (
    echo No se pudo actualizar pip. Intentando activar pip...
    "%APP_PY%" -m ensurepip --upgrade
)

"%APP_PY%" -m pip --version >nul 2>nul
if errorlevel 1 (
    if not "%APP_PY%"=="%PYTHON_EXE%" (
        echo pip no esta disponible en el entorno virtual. Se usara Python del usuario.
        set "APP_PY=%PYTHON_EXE%"
        "%APP_PY%" -m pip install --upgrade pip
    )
)

"%APP_PY%" -m pip --version >nul 2>nul
if errorlevel 1 (
    echo pip no esta disponible en este Python.
    echo Instala Python desde python.org marcando la opcion "Add Python to PATH".
    start https://www.python.org/downloads/
    pause
    exit /b 1
)

"%APP_PY%" -m pip install -r requirements.txt
if errorlevel 1 (
    echo No se pudieron instalar las dependencias.
    pause
    exit /b 1
)

"%APP_PY%" -c "import flask, pymupdf4llm, fitz" >nul 2>nul
if errorlevel 1 (
    echo Las dependencias no quedaron instaladas correctamente.
    pause
    exit /b 1
)

:deps_ready
echo.
echo Abriendo http://localhost:%PORT%
echo No cierres esta ventana mientras uses la app.
echo Presiona Ctrl+C para detener el servidor.
echo.

powershell -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 3; Start-Process 'http://localhost:%PORT%'" >nul 2>nul

"%APP_PY%" app.py %PORT%

pause
exit /b 0

:detect_python
set "PYTHON_CMD="
set "FIRST_PYTHON_CMD="

where py >nul 2>nul
if not errorlevel 1 (
    set "FIRST_PYTHON_CMD=py -3"
    py -3 -c "import flask, pymupdf4llm, fitz" >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_CMD=py -3"
        exit /b 0
    )
)

where python >nul 2>nul
if not errorlevel 1 (
    if "%FIRST_PYTHON_CMD%"=="" set "FIRST_PYTHON_CMD=python"
    python -c "import flask, pymupdf4llm, fitz" >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_CMD=python"
        exit /b 0
    )
)

where py >nul 2>nul
if not errorlevel 1 (
    py -3 -m pip --version >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_CMD=py -3"
        exit /b 0
    )
)

where python >nul 2>nul
if not errorlevel 1 (
    python -m pip --version >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_CMD=python"
        exit /b 0
    )
)

if not "%FIRST_PYTHON_CMD%"=="" (
    set "PYTHON_CMD=%FIRST_PYTHON_CMD%"
    exit /b 0
)

exit /b 0
