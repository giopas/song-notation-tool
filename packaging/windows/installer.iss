; Inno Setup script for Song Notation Tool (Windows).
; Build:  ISCC /DAppVersion=0.29.0 packaging\windows\installer.iss   (after pyinstaller)
; Per-user install (no administrator): the app folder stays writable, so
; "Update and restart" inside the app keeps working.
#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif

[Setup]
AppId={{7AE9406C-3B65-4457-8A2F-48E585D9BCF0}
AppName=Song Notation Tool
AppVersion={#AppVersion}
AppPublisher=giopas
AppPublisherURL=https://github.com/giopas/song-notation-tool
AppSupportURL=https://github.com/giopas/song-notation-tool/issues
DefaultDirName={autopf}\Song Notation Tool
DefaultGroupName=Song Notation Tool
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\..\release
OutputBaseFilename=Song-Notation-Tool-{#AppVersion}-windows-x64-setup
SetupIconFile=..\icons\icon.ico
UninstallDisplayIcon={app}\Song Notation Tool.exe
LicenseFile=..\..\LICENSE
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; Flags: unchecked

[Files]
Source: "..\..\dist\Song Notation Tool\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{autoprograms}\Song Notation Tool"; Filename: "{app}\Song Notation Tool.exe"
Name: "{autodesktop}\Song Notation Tool"; Filename: "{app}\Song Notation Tool.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Song Notation Tool.exe"; Description: "Start Song Notation Tool"; Flags: nowait postinstall skipifsilent
