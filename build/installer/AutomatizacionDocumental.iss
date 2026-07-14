#define MyAppName "AutomatizacionDocumental"
#define MyAppDisplayName "Automatizacion documental"
#ifndef MyAppVersion
  #define MyAppVersion "1.1.0"
#endif
#ifndef MySourceDir
  #define MySourceDir "dist\AutomatizacionDocumental"
#endif
#ifndef MyOutputDir
  #define MyOutputDir "dist\installer"
#endif
#ifndef MyUpdateSettingsFile
  #define MyUpdateSettingsFile "build\installer\update_settings.json"
#endif
#ifndef MySetupIconFile
  #define MySetupIconFile "assets\app_icon.ico"
#endif

[Setup]
AppId={{0BBD1CF4-9475-46BD-9D2E-F95B7E2EE10C}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppDisplayName}
DefaultDirName={autopf}\AutomatizacionDocumental
DefaultGroupName=AutomatizacionDocumental
DisableProgramGroupPage=yes
OutputDir={#MyOutputDir}
OutputBaseFilename=AutomatizacionDocumentalSetup_{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
SetupIconFile={#MySetupIconFile}
UninstallDisplayIcon={app}\AutomatizacionDocumental.exe

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear acceso directo en el escritorio"; GroupDescription: "Accesos directos:"; Flags: checkedonce

[Files]
Source: "{#MySourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#MyUpdateSettingsFile}"; DestDir: "{app}"; DestName: "update_settings.json"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppDisplayName}"; Filename: "{app}\AutomatizacionDocumental.exe"; WorkingDir: "{app}"
Name: "{autodesktop}\{#MyAppDisplayName}"; Filename: "{app}\AutomatizacionDocumental.exe"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\AutomatizacionDocumental.exe"; Description: "Abrir Automatizacion documental"; Flags: nowait postinstall skipifsilent
