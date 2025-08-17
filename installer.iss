; installer.iss
#define MyAppName "Cryo Label Studio"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Shannon Barrera"
#define MyAppExeName "Cryo Label Studio.exe"

[Setup]
AppId={{60a83225-065a-4dfa-adbd-f141e4f537a7}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableDirPage=no
DisableProgramGroupPage=yes
OutputDir=dist\installer
OutputBaseFilename=CryoLabelStudio-Setup
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=lowest
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
; Copy everything PyInstaller produced in dist\CryoPop Label Studio Lite\
Source: "dist\{#MyAppName}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{userdesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

; (Optional) Pre-create user data folders on install for first-run niceness
[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  userDataDir: string;
begin
  if CurStep = ssPostInstall then
  begin
    userDataDir := ExpandConstant('{localappdata}\'+ '{#MyAppName}');
    ForceDirectories(userDataDir + '\presets');
    ForceDirectories(userDataDir + '\logs');
  end;
end;
