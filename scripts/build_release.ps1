param(
    [string]$PythonPath = ".\.venv\Scripts\python.exe",
    [string]$TesseractPath = "$env:LOCALAPPDATA\Programs\Tesseract-OCR",
    [string]$InnoCompiler = ""
)

$ErrorActionPreference = "Stop"
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Set-Location -LiteralPath $projectRoot

function Assert-ChildPath([string]$Target) {
    $full = [System.IO.Path]::GetFullPath($Target)
    $rootWithSlash = $projectRoot.TrimEnd('\') + '\'
    if (-not $full.StartsWith($rootWithSlash, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Ruta fuera del proyecto: $full"
    }
    return $full
}

$resolvedPython = (Resolve-Path -LiteralPath $PythonPath).Path
$resolvedTesseract = (Resolve-Path -LiteralPath $TesseractPath).Path
$appVersion = (& $resolvedPython -c "from runtime_config import APP_VERSION; print(APP_VERSION)").Trim()
if (-not $appVersion) { throw "No se pudo determinar la versión de PDF2MD." }
$ocrRuntime = Assert-ChildPath (Join-Path $projectRoot "build\tesseract_runtime")
$distApp = Assert-ChildPath (Join-Path $projectRoot "dist\PDF2MD")
$releaseDir = Assert-ChildPath (Join-Path $projectRoot "release")

foreach ($language in @("spa", "eng", "osd")) {
    $trainedData = Join-Path $resolvedTesseract "tessdata\$language.traineddata"
    if (-not (Test-Path -LiteralPath $trainedData -PathType Leaf)) {
        throw "Falta el idioma OCR requerido: $trainedData"
    }
}

if (Test-Path -LiteralPath $ocrRuntime) {
    Remove-Item -LiteralPath $ocrRuntime -Recurse -Force
}
New-Item -ItemType Directory -Path (Join-Path $ocrRuntime "tessdata") -Force | Out-Null
foreach ($language in @("spa", "eng", "osd")) {
    Copy-Item -LiteralPath (Join-Path $resolvedTesseract "tessdata\$language.traineddata") -Destination (Join-Path $ocrRuntime "tessdata")
}

& $resolvedPython "scripts\create_brand_assets.py"
if ($LASTEXITCODE -ne 0) { throw "No se pudieron generar los recursos visuales." }

$sourceArchive = Assert-ChildPath (Join-Path $projectRoot "distribution\CODIGO_FUENTE_PDF2MD_$appVersion.zip")
if (Test-Path -LiteralPath $sourceArchive) {
    Remove-Item -LiteralPath $sourceArchive -Force
}
$sourceItems = @(
    "app.py",
    "launcher.py",
    "runtime_config.py",
    "conversion_formats.py",
    "requirements.txt",
    "requirements-build.txt",
    "README.md",
    "PDF2MD_Batch.spec",
    "installer",
    "scripts",
    "tests"
)
Compress-Archive -Path $sourceItems -DestinationPath $sourceArchive -CompressionLevel Optimal

if (Test-Path -LiteralPath $distApp) {
    Remove-Item -LiteralPath $distApp -Recurse -Force
}
& $resolvedPython -m PyInstaller --noconfirm --clean "PDF2MD_Batch.spec"
if ($LASTEXITCODE -ne 0) { throw "PyInstaller no pudo construir PDF2MD." }

$auditReport = Assert-ChildPath (Join-Path $projectRoot "distribution\AUDITORIA_TECNICA_PDF2MD_$appVersion.txt")
& $resolvedPython "scripts\audit_windows_bundle.py" --bundle $distApp --output $auditReport
if ($LASTEXITCODE -ne 0) { throw "La auditoría offline detectó archivos o dependencias faltantes." }

if (-not $InnoCompiler) {
    $candidates = @(
        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        "C:\Program Files\Inno Setup 6\ISCC.exe",
        (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"),
        (Join-Path $projectRoot ".tools\Inno Setup 6\ISCC.exe")
    )
    $InnoCompiler = $candidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
}
if (-not $InnoCompiler) {
    throw "No se encontró ISCC.exe. Instala Inno Setup 6 en el equipo de compilación."
}

New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null
Get-ChildItem -LiteralPath $releaseDir -File -ErrorAction SilentlyContinue |
    Where-Object {
        $_.Name -match '^Instalar_PDF2MD_.*\.(exe|sha256\.txt)$' -or
        $_.Name -match '^CODIGO_FUENTE_PDF2MD_.*\.zip$' -or
        $_.Name -match '^AUDITORIA_TECNICA_PDF2MD_.*\.txt$'
    } |
    Remove-Item -Force
& $InnoCompiler "installer\PDF2MD.iss"
if ($LASTEXITCODE -ne 0) { throw "Inno Setup no pudo construir el instalador." }

$installer = Join-Path $releaseDir "Instalar_PDF2MD_$appVersion.exe"
if (-not (Test-Path -LiteralPath $installer -PathType Leaf)) {
    throw "No se generó el instalador esperado: $installer"
}

Copy-Item -LiteralPath "output\pdf\Guia_rapida_PDF2MD.pdf" -Destination $releaseDir -Force
Copy-Item -LiteralPath "distribution\VERSION.txt" -Destination $releaseDir -Force
Copy-Item -LiteralPath "distribution\LICENCIAS_DE_TERCEROS.txt" -Destination $releaseDir -Force
Copy-Item -LiteralPath "distribution\CODIGO_FUENTE_PDF2MD_$appVersion.zip" -Destination $releaseDir -Force
Copy-Item -LiteralPath "distribution\AUDITORIA_TECNICA_PDF2MD_$appVersion.txt" -Destination $releaseDir -Force

$hash = (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash.ToLowerInvariant()
"$hash  Instalar_PDF2MD_$appVersion.exe" | Set-Content -LiteralPath (Join-Path $releaseDir "Instalar_PDF2MD_$appVersion.sha256.txt") -Encoding ascii

Write-Host "Entrega creada en: $releaseDir"
Write-Host "SHA-256: $hash"
