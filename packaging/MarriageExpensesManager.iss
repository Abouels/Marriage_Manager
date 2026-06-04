#define MyAppName "Marriage Expenses Manager"
#ifndef MyAppVersion
#define MyAppVersion "1.0.0"
#endif
#define MyAppPublisher "Eng. I. Abouelsaad"
#define MyAppExeName "MarriageExpensesManager.exe"

[Setup]
AppId={{6D8AA91B-7F31-42E0-9BC6-2E7E5A9A4E11}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\MarriageExpensesManager
DefaultGroupName={#MyAppName}
OutputDir=..\dist\installer
OutputBaseFilename=MarriageExpensesManager_Setup_{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest

[Files]
Source: "..\dist\MarriageExpensesManager\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs; Excludes: "portable.flag,app_data,app_data\*"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
