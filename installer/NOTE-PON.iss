#define MyAppName "NOTE-PON"
#define MyAppVersion "0.1.0"
#define MyAppPublisher "Tetsu Suzuki"
#define MyAppURL "https://github.com/pondashicom/note-pon"
#define MyAppExeName "NOTE-PON.exe"

[Setup]
AppId={{E8D56842-9101-4110-B547-F46C3864B51B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
LicenseFile=..\LICENSE
OutputDir=..\publish\installer
OutputBaseFilename=NOTE-PON-Setup-{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
SetupLogging=yes
CloseApplications=yes
CloseApplicationsFilter={#MyAppExeName}
RestartApplications=no
UninstallDisplayIcon={app}\{#MyAppExeName}
VersionInfoVersion=0.1.0.0
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription=NOTE-PON Installer
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}
VersionInfoCopyright=Copyright (c) 2026 {#MyAppPublisher}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "デスクトップにショートカットを作成"; GroupDescription: "追加のショートカット:"; Flags: unchecked

[Files]
Source: "..\bin\Release\net8.0-windows\NOTE-PON.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\bin\Release\net8.0-windows\NOTE-PON.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\bin\Release\net8.0-windows\NOTE-PON.deps.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\bin\Release\net8.0-windows\NOTE-PON.runtimeconfig.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE"; DestDir: "{app}"; DestName: "LICENSE.txt"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{#MyAppName}を起動"; Flags: nowait postinstall skipifsilent

[Code]
function HasDotNet8DesktopRuntime(): Boolean;
var
  RuntimeVersions: TArrayOfString;
  Index: Integer;
begin
  Result := False;

  if RegGetValueNames(
    HKLM32,
    'SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedfx\Microsoft.WindowsDesktop.App',
    RuntimeVersions) then
  begin
    for Index := 0 to GetArrayLength(RuntimeVersions) - 1 do
    begin
      if Pos('8.', RuntimeVersions[Index]) = 1 then
      begin
        Result := True;
        Exit;
      end;
    end;
  end;
end;

function InitializeSetup(): Boolean;
begin
  Result := True;

  if not HasDotNet8DesktopRuntime() then
  begin
    Result := MsgBox(
      'NOTE-PONの実行には.NET 8 Desktop RuntimeとWindowsデスクトップ版Microsoft PowerPointが別途必要です。' + #13#10 + #13#10 +
      'これらはインストーラーに含まれていません。インストールを続行しますか？',
      mbConfirmation,
      MB_YESNO) = IDYES;
  end;
end;
