#define MyAppName "NOTE-PON"
#define MyAppVersion "0.2.3"
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
MinVersion=10.0.26100
SetupLogging=yes
SetupIconFile=..\assets\note-pon.ico
CloseApplications=yes
CloseApplicationsFilter={#MyAppExeName}
RestartApplications=no
UninstallDisplayIcon={app}\{#MyAppExeName}
VersionInfoVersion=0.2.3.0
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
Source: "..\bin\Release\net10.0-windows\NOTE-PON.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\bin\Release\net10.0-windows\NOTE-PON.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\bin\Release\net10.0-windows\NOTE-PON.deps.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\bin\Release\net10.0-windows\NOTE-PON.runtimeconfig.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE"; DestDir: "{app}"; DestName: "LICENSE.txt"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{#MyAppName}を起動"; Flags: nowait postinstall skipifsilent

[Code]
function TryParseRuntimeVersion(VersionText: String; var Version: Int64): Boolean;
var
  DotPosition: Integer;
  MajorVersion: Integer;
  MinorVersion: Integer;
  PatchVersion: Integer;
  RemainingText: String;
begin
  Result := False;
  RemainingText := VersionText;

  DotPosition := Pos('.', RemainingText);
  if DotPosition = 0 then
    Exit;
  MajorVersion := StrToIntDef(Copy(RemainingText, 1, DotPosition - 1), -1);
  Delete(RemainingText, 1, DotPosition);

  DotPosition := Pos('.', RemainingText);
  if DotPosition = 0 then
    Exit;
  MinorVersion := StrToIntDef(Copy(RemainingText, 1, DotPosition - 1), -1);
  Delete(RemainingText, 1, DotPosition);
  PatchVersion := StrToIntDef(RemainingText, -1);

  if (MajorVersion < 0) or (MinorVersion < 0) or (PatchVersion < 0) then
    Exit;

  Version := PackVersionComponents(MajorVersion, MinorVersion, PatchVersion, 0);
  Result := True;
end;

function RegistryHasRequiredDesktopRuntime(RootKey: Integer): Boolean;
var
  RuntimeVersions: TArrayOfString;
  Index: Integer;
  InstalledVersion: Int64;
  MinimumVersion: Int64;
begin
  Result := False;
  MinimumVersion := PackVersionComponents(10, 0, 10, 0);

  if RegGetValueNames(
    RootKey,
    'SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedfx\Microsoft.WindowsDesktop.App',
    RuntimeVersions) then
  begin
    for Index := 0 to GetArrayLength(RuntimeVersions) - 1 do
    begin
      if TryParseRuntimeVersion(RuntimeVersions[Index], InstalledVersion) and
        (ComparePackedVersion(InstalledVersion, MinimumVersion) >= 0) then
      begin
        Result := True;
        Exit;
      end;
    end;
  end;
end;

function HasRequiredDotNet10DesktopRuntime(): Boolean;
begin
  Result :=
    RegistryHasRequiredDesktopRuntime(HKLM32) or
    RegistryHasRequiredDesktopRuntime(HKLM64);
end;

function InitializeSetup(): Boolean;
begin
  Result := HasRequiredDotNet10DesktopRuntime();
  if not Result then
  begin
    MsgBox(
      'NOTE-PONの実行には.NET 10 Desktop Runtime 10.0.10以降が必要です。' + #13#10 + #13#10 +
      '先に最新の.NET 10 Desktop Runtimeをインストールしてから、NOTE-PONのインストーラーを再実行してください。' + #13#10 +
      'https://dotnet.microsoft.com/download/dotnet/10.0',
      mbError,
      MB_OK);
  end;
end;
