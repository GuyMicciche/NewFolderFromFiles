; NewFolderFromFiles Installer Script
; Requires Inno Setup 6.x

#define MyAppName "New Folder From Files"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Your Name"
#define MyAppURL "https://github.com/yourusername/NewFolderFromFiles"
#define MyAppExeName "NewFolderFromFilesHotkey.exe"

[Setup]
AppId={{074AAE64-2F35-4E30-AF21-60A2F059E8F1}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=..\build
OutputBaseFilename=NewFolderFromFiles-Setup
Compression=lzma
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=admin
MinVersion=10.0

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "startupicon"; Description: "Start hotkey helper with Windows"; GroupDescription: "Additional options:"

[Files]
Source: "..\build\bin\Release\NewFolderFromFiles.dll"; DestDir: "{app}"; Flags: ignoreversion regserver 64bit
Source: "..\build\bin\Release\NewFolderFromFilesHotkey.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName} Hotkey Helper"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{userstartup}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: startupicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch hotkey helper"; Flags: nowait postinstall skipifsilent

[Code]
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  ResultCode: Integer;
begin
  if CurUninstallStep = usUninstall then
  begin
    // Kill hotkey helper if running
    Exec('taskkill.exe', '/F /IM NewFolderFromFilesHotkey.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    // Unregister DLL
    Exec('regsvr32.exe', '/u /s "' + ExpandConstant('{app}\NewFolderFromFiles.dll') + '"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  end;
end;
