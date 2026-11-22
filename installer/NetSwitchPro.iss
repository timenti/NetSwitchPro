#ifndef AppVersion
  #error AppVersion must be set from the published executable
#endif
#ifndef Payload
  #error Payload must be set to the Windows x64 executable
#endif
#define AppName "NetSwitch Pro"
#define AppExe "NetSwitchPro.exe"

[Setup]
AppId=NetSwitchPro.Timenti
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=Timenti
AppPublisherURL=https://timenti.com
AppUpdatesURL=https://github.com/timenti/NetSwitchPro/releases/latest
DefaultDirName={localappdata}\Programs\NetSwitchPro\Current
DisableDirPage=yes
UsePreviousAppDir=no
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern dynamic hidebevels
SetupIconFile=..\Assets\app.ico
UninstallDisplayIcon={app}\{#AppExe}
VersionInfoVersion={#AppVersion}
VersionInfoCompany=Timenti
VersionInfoProductName={#AppName}
OutputBaseFilename=NetSwitchPro-{#AppVersion}
CloseApplications=yes
CloseApplicationsFilter={#AppExe}
RestartApplications=no
Compression=lzma2
SolidCompression=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Files]
Source: "{#Payload}"; DestDir: "{app}"; DestName: "{#AppExe}"
Source: "..\LICENSE"; DestDir: "{app}"; DestName: "LICENSE.txt"

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"

[Run]
Filename: "{app}\{#AppExe}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent

[CustomMessages]
english.NewerInstalled=A newer version of NetSwitch Pro is installed.
russian.NewerInstalled=Уже установлена более новая версия NetSwitch Pro.

[Code]
const
  RunKey = 'Software\Microsoft\Windows\CurrentVersion\Run';
  RunName = 'NetSwitchPro';

function InstalledCommand: String;
begin
  Result := '"' + ExpandConstant('{app}\{#AppExe}') + '" --start-in-tray';
end;

function LegacyFile(Version, ExpectedHash: String): String;
var
  Candidate: String;
begin
  Result := '';
  Candidate := ExpandConstant('{localappdata}\Programs\NetSwitchPro\' + Version + '\{#AppExe}');
  if not FileExists(Candidate) then exit;
  try
    if CompareText(GetSHA256OfFile(Candidate), ExpectedHash) = 0 then
      Result := Candidate;
  except
    Log('Could not read legacy NetSwitch Pro executable: ' + Candidate);
  end;
end;

procedure HandleLegacyCopy(Version, ExpectedHash: String; RemoveCopy: Boolean);
var
  Candidate: String;
begin
  Candidate := LegacyFile(Version, ExpectedHash);
  if Candidate = '' then exit;
  if not RemoveCopy then begin
    RegisterExtraCloseApplicationsResource(Candidate);
  end else if DeleteFile(Candidate) then begin
    RemoveDir(ExtractFileDir(Candidate));
  end else begin
    Log('Legacy NetSwitch Pro executable is in use: ' + Candidate);
  end;
end;

procedure HandleLegacyCopies(RemoveCopies: Boolean);
begin
  HandleLegacyCopy('2.0.9', '313c2b990e0910ddff8458ebb867a04ec8f058bc83a5dec892ba93b6b6ebc414', RemoveCopies);
  HandleLegacyCopy('2.1.0', 'c1cde95045531c174bd2df51e8a4f9f2e84f5e8aae8f728bdf08800848db8b0b', RemoveCopies);
end;

procedure RegisterExtraCloseApplicationsResources;
begin
  HandleLegacyCopies(False);
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  CurrentVersion, InstallerVersion: Int64;
begin
  Result := '';
  if GetPackedVersion(ExpandConstant('{app}\{#AppExe}'), CurrentVersion) and
     GetPackedVersion(ExpandConstant('{srcexe}'), InstallerVersion) and
     (ComparePackedVersion(CurrentVersion, InstallerVersion) > 0) then
    Result := CustomMessage('NewerInstalled');
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  Existing: String;
begin
  if CurStep <> ssPostInstall then exit;
  if RegQueryStringValue(HKCU, RunKey, RunName, Existing) and
     (CompareText(Existing, InstalledCommand) <> 0) and
     not RegWriteStringValue(HKCU, RunKey, RunName, InstalledCommand) then begin
    Log('Could not migrate the NetSwitch Pro startup registration.');
    exit;
  end;
  HandleLegacyCopies(True);
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  Existing: String;
begin
  if (CurUninstallStep = usUninstall) and RegQueryStringValue(HKCU, RunKey, RunName, Existing) and
     (CompareText(Existing, InstalledCommand) = 0) then
    RegDeleteValue(HKCU, RunKey, RunName);
end;
