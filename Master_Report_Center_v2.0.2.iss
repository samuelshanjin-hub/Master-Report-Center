#define MyAppName "Master Report Center"
#define MyAppVersion "2.0.2"
#define MyAppPublisher "Golden Fir Solutions"
#define MyAppExeName "Master Report Center.exe"

[Setup]
AppId={{D873C4A8-3185-4F7F-9C64-54CF17D15D02}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}

DefaultDirName={autopf}\Golden Fir Solutions\Master Report Center
DefaultGroupName=Golden Fir Solutions\Master Report Center

DisableProgramGroupPage=yes
PrivilegesRequired=admin

ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

OutputDir=installer_output
OutputBaseFilename=Master_Report_Center_v2.0.2_Setup

Compression=lzma2
SolidCompression=yes
WizardStyle=modern

UninstallDisplayName={#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}

CloseApplications=yes
RestartApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "dist\Master Report Center\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Master Report Center"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\Master Report Center"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch Master Report Center"; Flags: nowait postinstall skipifsilent