#ifndef AppVersion
  #error AppVersion must be set from the published executable
#endif
#ifndef Payload
  #error Payload must be set to the Windows x64 executable
#endif
#ifndef DisplayVersion
  #define DisplayVersion AppVersion
#endif
#define AppName "NetSwitch Pro"
#define AppExe "NetSwitchPro.exe"

[Setup]
AppId=NetSwitchPro.Timenti
AppName={#AppName}
AppVersion={#DisplayVersion}
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
WizardStyle=modern dark hidebevels
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
Source: "..\Assets\app.ico"; DestName: "progress-icon.ico"; Flags: dontcopy noencryption
Source: "{#Payload}"; DestDir: "{app}"; DestName: "{#AppExe}"; BeforeInstall: SetInstallStage('ProgressExtracting')
Source: "..\LICENSE"; DestDir: "{app}"; DestName: "LICENSE.txt"; BeforeInstall: SetInstallStage('ProgressExtracting')

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"; BeforeInstall: SetInstallStage('ProgressShortcuts')

[Run]
Filename: "{app}\{#AppExe}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent

[CustomMessages]
english.NewerInstalled=A newer version of NetSwitch Pro is installed.
russian.NewerInstalled=Уже установлена более новая версия NetSwitch Pro.
english.ProgressInstall=Installation
russian.ProgressInstall=Установка
english.ProgressUpdate=Update
russian.ProgressUpdate=Обновление
english.ProgressPreparing=Preparing installation…
russian.ProgressPreparing=Подготавливаем установку…
english.ProgressExtracting=Extracting and installing…
russian.ProgressExtracting=Распаковываем и устанавливаем…
english.ProgressShortcuts=Setting up shortcuts…
russian.ProgressShortcuts=Настраиваем ярлыки…
english.ProgressFinalizing=Finishing installation…
russian.ProgressFinalizing=Завершаем установку…

[Code]
function BitBlt(Destination: HDC; X, Y, Width, Height: Integer; Source: HDC; SourceX, SourceY: Integer; Operation: Cardinal): Boolean;
external 'BitBlt@gdi32.dll stdcall';

const
  RunKey = 'Software\Microsoft\Windows\CurrentVersion\Run';
  RunName = 'NetSwitchPro';
  ProgressBackground = $161616;
  ProgressTrack = $3F3F3F;
  ProgressBadge = $252525;
  ProgressInk = $F6F6F6;
  ProgressMuted = $ADADAD;
  CompactWidth = 296;
  CompactHeight = 137;
  CopyPixels = $00CC0020;

var
  ProgressPanel: TPanel;
  ProgressImage: TBitmapImage;
  ProgressIcon: TBitmapImage;
  ProgressCaption, ProgressStatus: String;
  ProgressCurrent: Integer;
  OriginalWidth, OriginalHeight, OriginalLeft, OriginalTop: Integer;

procedure WriteProgressText(Surface: TCanvas; const Caption: String; Top, Height, FontPixels: Integer; Color: TColor; RightAligned: Boolean);
var
  TextWidth, Left: Integer;
begin
  Surface.Font.Name := 'Segoe UI';
  if FontPixels = 25 then Surface.Font.Name := 'Segoe UI Semibold';
  Surface.Font.Height := -ScaleY(FontPixels);
  Surface.Font.Color := Color;
  Surface.Brush.Style := bsClear;
  TextWidth := Surface.TextWidth(Caption);
  if RightAligned then Left := ProgressImage.Width - ScaleX(17) - TextWidth
  else Left := (ProgressImage.Width - TextWidth) div 2;
  Surface.TextOut(Left, ScaleY(Top) + (ScaleY(Height) - Surface.TextHeight(Caption)) div 2, Caption);
  Surface.Brush.Style := bsSolid;
end;

procedure DrawProgressBrand(Surface: TCanvas);
var
  BadgeLeft, BadgeRight, BadgeTop, BadgeBottom, VersionWidth: Integer;
begin
  if not ProgressIcon.PngImage.Empty then
    BitBlt(Surface.Handle, ScaleX(17), ScaleY(13), ProgressIcon.PngImage.Width, ProgressIcon.PngImage.Height,
      ProgressIcon.PngImage.Canvas.Handle, 0, 0, CopyPixels)
  else if not ProgressIcon.Bitmap.Empty then
    BitBlt(Surface.Handle, ScaleX(17), ScaleY(13), ProgressIcon.Bitmap.Width, ProgressIcon.Bitmap.Height,
      ProgressIcon.Bitmap.Canvas.Handle, 0, 0, CopyPixels);

  Surface.Font.Name := 'Segoe UI Semibold';
  Surface.Font.Height := -ScaleY(16);
  Surface.Font.Color := ProgressInk;
  Surface.Brush.Style := bsClear;
  Surface.TextOut(ScaleX(53), ScaleY(11) + (ScaleY(33) - Surface.TextHeight('{#AppName}')) div 2, '{#AppName}');

  Surface.Font.Name := 'Segoe UI';
  Surface.Font.Height := -ScaleY(13);
  VersionWidth := Surface.TextWidth('{#DisplayVersion}');
  BadgeRight := ProgressImage.Width - ScaleX(17);
  BadgeLeft := BadgeRight - VersionWidth - ScaleX(16);
  BadgeTop := ScaleY(17);
  BadgeBottom := BadgeTop + ScaleY(21);
  Surface.Brush.Style := bsSolid;
  Surface.Brush.Color := ProgressBadge;
  Surface.Pen.Style := psSolid;
  Surface.Pen.Color := ProgressTrack;
  Surface.RoundRect(BadgeLeft, BadgeTop, BadgeRight, BadgeBottom, ScaleX(14), ScaleY(14));
  Surface.Brush.Style := bsClear;
  Surface.TextOut(BadgeLeft + ScaleX(8), BadgeTop + (ScaleY(21) - Surface.TextHeight('{#DisplayVersion}')) div 2, '{#DisplayVersion}');
  Surface.Brush.Style := bsSolid;
end;

procedure DrawProgress(Percent: Integer);
var
  Surface: TCanvas;
  Bounds: TRect;
  RailLeft, RailTop, RailWidth, Radius, FilledWidth: Integer;
begin
  Bounds.Left := 0;
  Bounds.Top := 0;
  Bounds.Right := ProgressImage.Width;
  Bounds.Bottom := ProgressImage.Height;
  Surface := ProgressImage.Bitmap.Canvas;
  Surface.Brush.Color := ProgressBackground;
  Surface.FillRect(Bounds);

  DrawProgressBrand(Surface);
  WriteProgressText(Surface, ProgressCaption, 44, 35, 25, ProgressInk, False);
  WriteProgressText(Surface, ProgressStatus, 79, 22, 13, ProgressMuted, False);
  WriteProgressText(Surface, IntToStr(Percent) + '%', 101, 22, 13, ProgressInk, True);

  RailLeft := ScaleX(17);
  RailTop := ScaleY(109);
  RailWidth := ProgressImage.Width - ScaleX(42) - Surface.TextWidth('100%');
  Radius := ScaleY(5);
  Surface.Pen.Style := psClear;
  Surface.Brush.Color := ProgressTrack;
  Surface.RoundRect(RailLeft, RailTop, RailLeft + RailWidth, RailTop + Radius, Radius, Radius);
  FilledWidth := (RailWidth * Percent) div 100;
  if FilledWidth > 0 then begin
    Surface.Brush.Color := ProgressInk;
    Surface.RoundRect(RailLeft, RailTop, RailLeft + FilledWidth, RailTop + Radius, Radius, Radius);
  end;
  ProgressImage.Invalidate;
end;

procedure SetInstallStage(const MessageKey: String);
begin
  ProgressStatus := CustomMessage(MessageKey);
  if ProgressPanel.Visible then begin
    DrawProgress(ProgressCurrent);
    ProgressImage.Update;
  end;
end;

procedure InitializeWizard;
begin
  ProgressPanel := TPanel.Create(WizardForm);
  ProgressPanel.Parent := WizardForm;
  ProgressPanel.Visible := False;
  ProgressPanel.Align := alClient;
  ProgressPanel.BevelOuter := bvNone;
  ProgressPanel.StyleElements := [];
  ProgressPanel.ParentBackground := False;
  ProgressPanel.Color := ProgressBackground;
  ProgressCaption := CustomMessage('ProgressInstall');
  ProgressStatus := CustomMessage('ProgressPreparing');

  ProgressIcon := TBitmapImage.Create(ProgressPanel);
  ProgressIcon.SetBounds(0, 0, ScaleX(28), ScaleY(28));
  try
    ExtractTemporaryFile('progress-icon.ico');
    if not InitializeBitmapImageFromIcon(ProgressIcon, ExpandConstant('{tmp}\progress-icon.ico'), ProgressBackground, []) then
      Log('Could not load the NetSwitch Pro icon for the update screen.');
  except
    Log('Could not extract the NetSwitch Pro icon for the update screen.');
  end;

  ProgressImage := TBitmapImage.Create(ProgressPanel);
  ProgressImage.Parent := ProgressPanel;
  ProgressImage.AutoSize := False;
  ProgressImage.SetBounds(0, 0, ScaleX(CompactWidth), ScaleY(CompactHeight));
  ProgressImage.Bitmap.Width := ProgressImage.Width;
  ProgressImage.Bitmap.Height := ProgressImage.Height;
  DrawProgress(0);
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if ProgressPanel = nil then exit;
  if CurPageID = wpInstalling then begin
    OriginalWidth := WizardForm.ClientWidth;
    OriginalHeight := WizardForm.ClientHeight;
    OriginalLeft := WizardForm.Left;
    OriginalTop := WizardForm.Top;
    WizardForm.ClientWidth := ScaleX(CompactWidth);
    WizardForm.ClientHeight := ScaleY(CompactHeight);
    WizardForm.Left := OriginalLeft + (OriginalWidth - WizardForm.ClientWidth) div 2;
    WizardForm.Top := OriginalTop + (OriginalHeight - WizardForm.ClientHeight) div 2;
    if FileExists(AddBackslash(WizardDirValue) + '{#AppExe}') then
      ProgressCaption := CustomMessage('ProgressUpdate');
    DrawProgress(0);
    ProgressPanel.Visible := True;
    ProgressPanel.BringToFront;
  end else if ProgressPanel.Visible then begin
    ProgressPanel.Visible := False;
    WizardForm.ClientWidth := OriginalWidth;
    WizardForm.ClientHeight := OriginalHeight;
    WizardForm.Left := OriginalLeft;
    WizardForm.Top := OriginalTop;
  end;
end;

procedure CurInstallProgressChanged(CurProgress, MaxProgress: Integer);
var
  Percent: Integer;
begin
  if MaxProgress <= 0 then exit;
  Percent := Round(99 * (CurProgress / MaxProgress));
  if Percent > ProgressCurrent then begin
    ProgressCurrent := Percent;
    DrawProgress(ProgressCurrent);
  end;
end;

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
  SetInstallStage('ProgressFinalizing');
  if RegQueryStringValue(HKCU, RunKey, RunName, Existing) and
     (CompareText(Existing, InstalledCommand) <> 0) and
     not RegWriteStringValue(HKCU, RunKey, RunName, InstalledCommand) then begin
    Log('Could not migrate the NetSwitch Pro startup registration.');
    exit;
  end;
  HandleLegacyCopies(True);
  ProgressCurrent := 100;
  DrawProgress(ProgressCurrent);
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  Existing: String;
begin
  if (CurUninstallStep = usUninstall) and RegQueryStringValue(HKCU, RunKey, RunName, Existing) and
     (CompareText(Existing, InstalledCommand) = 0) then
    RegDeleteValue(HKCU, RunKey, RunName);
end;
