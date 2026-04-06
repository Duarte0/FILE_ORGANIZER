#define MyAppName "Distribuidor de Arquivos"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "DistribuidorArquivos"
#define MyAppExeName "DistribuidorArquivos.exe"

#ifndef MyAppSource
  #define MyAppSource "release\\app"
#endif

[Setup]
AppId={{6C3285A4-7134-4C60-9F8C-EE7D2D11C845}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\DistribuidorArquivos
DefaultGroupName=Distribuidor de Arquivos
DisableProgramGroupPage=yes
OutputDir=release
OutputBaseFilename=DistribuidorArquivos_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "portuguesebrazil"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Dirs]
Name: "{app}\logs"
Name: "{app}\relatorios"

[Files]
Source: "{#MyAppSource}\DistribuidorArquivos.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#MyAppSource}\config.env"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#MyAppSource}\regras.xlsx"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Distribuidor de Arquivos"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\Distribuidor de Arquivos"; Filename: "{app}\{#MyAppExeName}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Executar Distribuidor de Arquivos"; Flags: nowait postinstall skipifsilent