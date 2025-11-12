; Script de Inno Setup - Reintegros Nomina
#define MyAppName "Generador de reintergos"
#define MyAppVersion "1.0"
#define MyAppPublisher "Coordinación de Servicios Educativos Colima"
#define MyAppURL "https://github.com/MaxPena7/py_reintegros.git"
#define MyAppExeName "Reintegros_nomina.exe"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-1234-567890ABCDEF}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
LicenseFile=licencia.txt
OutputDir=Output
OutputBaseFilename=Instalador_Reintegros_Colima
SetupIconFile=reintegro_icono.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Aplicación principal
Source: "dist\Reintegros_nomina.exe"; DestDir: "{app}"; Flags: ignoreversion
; Archivos de recursos
Source: "plantilla.html"; DestDir: "{app}"; Flags: ignoreversion
Source: "fondo_reintegro.pdf"; DestDir: "{app}"; Flags: ignoreversion
Source: "reintegro_icono.ico"; DestDir: "{app}"; Flags: ignoreversion
; WKHTMLTOPDF incluido
Source: "wkhtmltopdf\*"; DestDir: "{app}\wkhtmltopdf"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\reintegro_icono.ico"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; IconFilename: "{app}\reintegro_icono.ico"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent