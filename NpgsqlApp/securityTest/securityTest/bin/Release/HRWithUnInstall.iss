; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "HR Apps"
#define MyAppVersion "1.0.1.0"
#define MyAppPublisher "DJ Soft Copyright �  2011"
#define MyAppExeName "HR.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{50D223FA-7B3D-48D8-A766-E9C021F2711F}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={pf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputBaseFilename=setup
Compression=lzma
SolidCompression=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 0,6.1

[Files]
Source: "D:\VB2010Work\NewAppWithDbsource\NpgsqlApp\securityTest\securityTest\bin\Release\HR.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\VB2010Work\NewAppWithDbsource\NpgsqlApp\securityTest\securityTest\bin\Release\DJLib.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\VB2010Work\NewAppWithDbsource\NpgsqlApp\securityTest\securityTest\bin\Release\DJLib.pdb"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\VB2010Work\NewAppWithDbsource\NpgsqlApp\securityTest\securityTest\bin\Release\DJLib.xml"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\VB2010Work\NewAppWithDbsource\NpgsqlApp\securityTest\securityTest\bin\Release\HR.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\VB2010Work\NewAppWithDbsource\NpgsqlApp\securityTest\securityTest\bin\Release\HR.exe.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\VB2010Work\NewAppWithDbsource\NpgsqlApp\securityTest\securityTest\bin\Release\HR.pdb"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\VB2010Work\NewAppWithDbsource\NpgsqlApp\securityTest\securityTest\bin\Release\HR.xml"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\VB2010Work\NewAppWithDbsource\NpgsqlApp\securityTest\securityTest\bin\Release\Mono.Security.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\VB2010Work\NewAppWithDbsource\NpgsqlApp\securityTest\securityTest\bin\Release\de"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "D:\VB2010Work\NewAppWithDbsource\NpgsqlApp\securityTest\securityTest\bin\Release\es"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "D:\VB2010Work\NewAppWithDbsource\NpgsqlApp\securityTest\securityTest\bin\Release\fi"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "D:\VB2010Work\NewAppWithDbsource\NpgsqlApp\securityTest\securityTest\bin\Release\zh-CN"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "D:\VB2010Work\NewAppWithDbsource\NpgsqlApp\securityTest\securityTest\bin\Release\fr"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "D:\VB2010Work\NewAppWithDbsource\NpgsqlApp\securityTest\securityTest\bin\Release\ja"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, "&", "&&")}}"; Flags: nowait postinstall skipifsilent
