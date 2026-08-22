[CmdletBinding()]
param(
    [switch]$NoLaunch
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$appName = "PDF2MD"
$version = "1.0.3"
$repository = "alejo961/pdf2md-batch"
$installerName = "Instalar_PDF2MD_$version.exe"
$expectedSha256 = "cf3fd62a42fe68f8906d447a2bb087b70ba692bf384cba988c078b1523eeee79"
$downloadUrl = "https://github.com/$repository/releases/download/v$version/$installerName"
$installDirectory = Join-Path $env:LOCALAPPDATA "Programs\PDF2MD"
$installedExecutable = Join-Path $installDirectory "PDF2MD.exe"

if (-not [Environment]::Is64BitOperatingSystem) {
    throw "PDF2MD requiere Windows 10 u 11 de 64 bits."
}

$temporaryRoot = [IO.Path]::GetTempPath().TrimEnd('\')
$workDirectory = Join-Path $temporaryRoot ("PDF2MD-install-" + [Guid]::NewGuid().ToString("N"))
$installerPath = Join-Path $workDirectory $installerName

try {
    New-Item -ItemType Directory -Path $workDirectory | Out-Null
    Write-Host "Descargando PDF2MD $version desde GitHub..."
    Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath -UseBasicParsing

    $actualSha256 = (Get-FileHash -LiteralPath $installerPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actualSha256 -ne $expectedSha256) {
        throw "La descarga no superó la verificación de seguridad SHA-256. No se ejecutará."
    }

    # The hash is verified before removing the internet-zone marker.
    Unblock-File -LiteralPath $installerPath
    Write-Host "Instalando PDF2MD para el usuario actual..."
    $installerArguments = @("/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/SP-")
    $installation = Start-Process -FilePath $installerPath -ArgumentList $installerArguments -Wait -PassThru
    if ($installation.ExitCode -ne 0) {
        throw "El instalador terminó con el código $($installation.ExitCode)."
    }
    if (-not (Test-Path -LiteralPath $installedExecutable -PathType Leaf)) {
        throw "La instalación terminó, pero no se encontró $installedExecutable."
    }

    # Inno Setup normally creates this shortcut. This fallback guarantees it
    # for automated installations and redirected OneDrive desktops.
    $desktopDirectory = [Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)
    $shortcutPath = Join-Path $desktopDirectory "PDF2MD.lnk"
    if (-not (Test-Path -LiteralPath $shortcutPath -PathType Leaf)) {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $installedExecutable
        $shortcut.WorkingDirectory = $installDirectory
        $shortcut.IconLocation = "$installedExecutable,0"
        $shortcut.Description = "Abrir PDF2MD en el navegador"
        $shortcut.Save()
    }
    if (-not (Test-Path -LiteralPath $shortcutPath -PathType Leaf)) {
        throw "PDF2MD se instaló, pero no fue posible crear el acceso directo del escritorio."
    }

    Write-Host "PDF2MD $version instalado correctamente."
    Write-Host "Acceso directo: $shortcutPath"
    if (-not $NoLaunch) {
        Start-Process -FilePath $installedExecutable | Out-Null
    }
}
finally {
    if (Test-Path -LiteralPath $installerPath -PathType Leaf) {
        Remove-Item -LiteralPath $installerPath -Force
    }
    if (Test-Path -LiteralPath $workDirectory -PathType Container) {
        Remove-Item -LiteralPath $workDirectory -Force
    }
}
