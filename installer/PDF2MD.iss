#define MyAppName "PDF2MD"
#define MyAppVersion "1.0.3"
#define MyAppPublisher "PDF2MD"
#define MyAppExeName "PDF2MD.exe"

[Setup]
AppId={{3E329218-03CA-4DA4-A6DA-3E4709D61322}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
VersionInfoVersion={#MyAppVersion}.0
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription=Instalador autocontenido de PDF2MD
VersionInfoProductName={#MyAppName}
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\release
OutputBaseFilename=Instalar_PDF2MD_{#MyAppVersion}
SetupIconFile=..\assets\pdf2md.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2/fast
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
RestartApplications=no
AppMutex=Local\PDF2MD_Desktop_Application
LicenseFile=..\distribution\LICENCIAS_DE_TERCEROS.txt
MinVersion=10.0

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Files]
Source: "..\dist\PDF2MD\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\output\pdf\Guia_rapida_PDF2MD.pdf"; DestDir: "{app}\Ayuda"; Flags: ignoreversion
Source: "..\distribution\LICENCIAS_DE_TERCEROS.txt"; DestDir: "{app}\Ayuda"; Flags: ignoreversion
Source: "..\distribution\VERSION.txt"; DestDir: "{app}\Ayuda"; Flags: ignoreversion
Source: "..\distribution\CODIGO_FUENTE_PDF2MD_1.0.3.zip"; DestDir: "{app}\Ayuda"; Flags: ignoreversion
Source: "..\distribution\AUDITORIA_TECNICA_PDF2MD_1.0.3.txt"; DestDir: "{app}\Ayuda"; Flags: ignoreversion

[Icons]
Name: "{autodesktop}\PDF2MD"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\{#MyAppExeName}"
Name: "{group}\PDF2MD"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\{#MyAppExeName}"
Name: "{group}\Guia rapida de PDF2MD"; Filename: "{app}\Ayuda\Guia_rapida_PDF2MD.pdf"
Name: "{group}\Desinstalar PDF2MD"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir PDF2MD"; Flags: nowait postinstall skipifsilent
