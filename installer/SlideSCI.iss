#define AppName "SlideBridge Office"
#define AddInKey "SlideSCICompat"
#define VstoManifest "SlideSCICompat.vsto"
#define SolutionId "{EDE5B327-B8B0-4044-9237-768C42B63E3E}"
#define InclusionId "{80B4B921-FA89-4AAE-8146-62F13CCC93E4}"
#define AppPublisher "jacywallny"
#define AppVersion GetEnv("SLIDESCI_VERSION")
#define PublishDir GetEnv("SLIDESCI_PUBLISH_DIR")
#define DistDir GetEnv("SLIDESCI_DIST_DIR")

#if AppVersion == ""
  #define AppVersion "1.0.0.0"
#endif

#if PublishDir == ""
  #define PublishDir "..\artifacts\publish"
#endif

#if DistDir == ""
  #define DistDir "..\artifacts\dist"
#endif

[Setup]
AppId={{0E5CB4DB-3A3D-4AB8-B8CC-561DD944451A}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName=D:\SlideSCI_WPS_PowerPoint_Compat
CreateAppDir=yes
Uninstallable=yes
PrivilegesRequired=lowest
OutputDir={#DistDir}
OutputBaseFilename=SlideBridge_Office_WPS_PowerPoint_v{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
DisableProgramGroupPage=yes
DisableReadyPage=no
DisableFinishedPage=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "{#PublishDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#AddInKey}"; ValueType: string; ValueName: "Description"; ValueData: "SlideBridge Office"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#AddInKey}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "SlideBridge"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#AddInKey}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#AddInKey}"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\{#VstoManifest}|vstolocal"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Kingsoft\Office\WPP\AddinsWL"; ValueType: string; ValueName: "{#AddInKey}"; ValueData: ""; Flags: uninsdeletevalue

[Code]
function FileUrl(Path: string): string;
begin
  StringChangeEx(Path, '\', '/', True);
  Result := 'file:///' + Path;
end;

function ExtractPublicKey(ManifestPath: string): string;
var
  Text: AnsiString;
  StartPos: Integer;
  EndPos: Integer;
  StartTag: string;
  EndTag: string;
begin
  Result := '';
  StartTag := '<RSAKeyValue>';
  EndTag := '</RSAKeyValue>';

  if LoadStringFromFile(ManifestPath, Text) then
  begin
    StartPos := Pos(StartTag, String(Text));
    if StartPos > 0 then
    begin
      EndPos := Pos(EndTag, String(Text));
      if EndPos > StartPos then
      begin
        Result := Copy(String(Text), StartPos, EndPos - StartPos + Length(EndTag));
      end;
    end;
  end;
end;

function ContainsSlideSci(Text: string): Boolean;
begin
  Result := Pos('SlideSCI', Text) > 0;
end;

procedure StopOfficeHosts;
var
  ResultCode: Integer;
begin
  Exec(
    ExpandConstant('{cmd}'),
    '/C powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "Stop-Process -Name POWERPNT,WPP,WPS -Force -ErrorAction SilentlyContinue"',
    '',
    SW_HIDE,
    ewWaitUntilTerminated,
    ResultCode
  );
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  StopOfficeHosts;
  Result := '';
end;

procedure CleanupStaleVstoRegistry(CurrentManifestUrl: string);
var
  InclusionRoot: string;
  MetaRoot: string;
  Names: TArrayOfString;
  ValueNames: TArrayOfString;
  I: Integer;
  Url: string;
  SolutionId: string;
  AddInName: string;
  FriendlyName: string;
  Description: string;
  KeyPath: string;
begin
  InclusionRoot := 'Software\Microsoft\VSTO\Security\Inclusion';
  MetaRoot := 'Software\Microsoft\VSTO\SolutionMetadata';

  if RegGetSubkeyNames(HKCU, InclusionRoot, Names) then
  begin
    for I := 0 to GetArrayLength(Names) - 1 do
    begin
      KeyPath := InclusionRoot + '\' + Names[I];
      if RegQueryStringValue(HKCU, KeyPath, 'Url', Url) then
      begin
        if ContainsSlideSci(Url) and (CompareText(Url, CurrentManifestUrl) <> 0) then
        begin
          RegDeleteKeyIncludingSubkeys(HKCU, KeyPath);
        end;
      end;
    end;
  end;

  if RegGetValueNames(HKCU, MetaRoot, ValueNames) then
  begin
    for I := 0 to GetArrayLength(ValueNames) - 1 do
    begin
      if ContainsSlideSci(ValueNames[I]) and (CompareText(ValueNames[I], CurrentManifestUrl) <> 0) then
      begin
        if RegQueryStringValue(HKCU, MetaRoot, ValueNames[I], SolutionId) then
        begin
          if CompareText(SolutionId, '{#SolutionId}') <> 0 then
          begin
            RegDeleteKeyIncludingSubkeys(HKCU, MetaRoot + '\' + SolutionId);
          end;
        end;

        RegDeleteValue(HKCU, MetaRoot, ValueNames[I]);
      end;
    end;
  end;

  if RegGetSubkeyNames(HKCU, MetaRoot, Names) then
  begin
    for I := 0 to GetArrayLength(Names) - 1 do
    begin
      if CompareText(Names[I], '{#SolutionId}') <> 0 then
      begin
        KeyPath := MetaRoot + '\' + Names[I];
        AddInName := '';
        FriendlyName := '';
        Description := '';
        RegQueryStringValue(HKCU, KeyPath, 'addInName', AddInName);
        RegQueryStringValue(HKCU, KeyPath, 'friendlyName', FriendlyName);
        RegQueryStringValue(HKCU, KeyPath, 'description', Description);

        if ContainsSlideSci(AddInName) or ContainsSlideSci(FriendlyName) or ContainsSlideSci(Description) then
        begin
          RegDeleteKeyIncludingSubkeys(HKCU, KeyPath);
        end;
      end;
    end;
  end;
end;

procedure CleanupStalePowerPointAddinPaths(CurrentManifestUrl: string);
var
  AddinsRoot: string;
  Names: TArrayOfString;
  I: Integer;
  KeyPath: string;
  ManifestValue: string;
  Description: string;
  FriendlyName: string;
  KeepCurrent: Boolean;
begin
  AddinsRoot := 'Software\Microsoft\Office\PowerPoint\Addins';

  if RegGetSubkeyNames(HKCU, AddinsRoot, Names) then
  begin
    for I := 0 to GetArrayLength(Names) - 1 do
    begin
      KeyPath := AddinsRoot + '\' + Names[I];
      ManifestValue := '';
      Description := '';
      FriendlyName := '';
      RegQueryStringValue(HKCU, KeyPath, 'Manifest', ManifestValue);
      RegQueryStringValue(HKCU, KeyPath, 'Description', Description);
      RegQueryStringValue(HKCU, KeyPath, 'FriendlyName', FriendlyName);

      KeepCurrent := False;
      if ManifestValue <> '' then
      begin
        KeepCurrent :=
          (CompareText(ManifestValue, CurrentManifestUrl) = 0) or
          (CompareText(ManifestValue, CurrentManifestUrl + '|vstolocal') = 0);
      end;

      if
        (
          ContainsSlideSci(Names[I]) or
          ContainsSlideSci(ManifestValue) or
          ContainsSlideSci(Description) or
          ContainsSlideSci(FriendlyName)
        )
        and (not KeepCurrent)
      then
      begin
        RegDeleteKeyIncludingSubkeys(HKCU, KeyPath);
      end;
    end;
  end;
end;

procedure CleanupWpsAddinState;
var
  LegacyAddinsRoot: string;
  Names: TArrayOfString;
  I: Integer;
  KeyPath: string;
begin
  RegDeleteValue(HKCU, 'Software\Kingsoft\Office\WPP\AddinsWL', 'SlideSCI');
  RegDeleteValue(HKCU, 'Software\Kingsoft\Office\WPP\AddinsWL', '{#AddInKey}');
  RegDeleteValue(HKCU, 'Software\Kingsoft\Office\WPP\AddinsBL', 'SlideSCI');
  RegDeleteValue(HKCU, 'Software\Kingsoft\Office\WPP\AddinsBL', '{#AddInKey}');
  RegDeleteValue(HKCU, 'Software\Kingsoft\Office\WPP\AddinsCL', 'SlideSCI');
  RegDeleteValue(HKCU, 'Software\Kingsoft\Office\WPP\AddinsCL', '{#AddInKey}');

  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Kingsoft\Office\6.0\WPP\AddIns\SlideSCI');
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Kingsoft\Office\6.0\WPP\AddIns\{#AddInKey}');

  LegacyAddinsRoot := 'Software\Kingsoft\Office\6.0\WPP\AddIns';
  if RegGetSubkeyNames(HKCU, LegacyAddinsRoot, Names) then
  begin
    for I := 0 to GetArrayLength(Names) - 1 do
    begin
      if ContainsSlideSci(Names[I]) then
      begin
        KeyPath := LegacyAddinsRoot + '\' + Names[I];
        RegDeleteKeyIncludingSubkeys(HKCU, KeyPath);
      end;
    end;
  end;
end;

procedure RegisterVstoTrust;
var
  ManifestUrl: string;
  ManifestPath: string;
  PublicKey: string;
  MetaRoot: string;
  InclusionKey: string;
begin
  ManifestPath := ExpandConstant('{app}\{#VstoManifest}');
  ManifestUrl := FileUrl(ManifestPath);
  PublicKey := ExtractPublicKey(ManifestPath);
  MetaRoot := 'Software\Microsoft\VSTO\SolutionMetadata';
  InclusionKey := 'Software\Microsoft\VSTO\Security\Inclusion\{#InclusionId}';

  CleanupStaleVstoRegistry(ManifestUrl);
  CleanupStalePowerPointAddinPaths(ManifestUrl);
  CleanupWpsAddinState;

  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Microsoft\Office\PowerPoint\Addins\SlideSCI');
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Microsoft\Office\PowerPoint\Addins\{#AddInKey}');
  RegDeleteValue(HKCU, MetaRoot, 'file:///D:/codex_cli/SlideSCI/SlideSCI/bin/Release/SlideSCICompat.vsto');
  RegDeleteValue(HKCU, MetaRoot, 'file:///D:/SlideSCI_wps/SlideSCI/bin/Release/SlideSCICompat.vsto');
  RegDeleteKeyIncludingSubkeys(HKCU, InclusionKey);

  RegWriteStringValue(HKCU, 'Software\Microsoft\Office\PowerPoint\Addins\{#AddInKey}', 'Description', 'SlideBridge Office');
  RegWriteStringValue(HKCU, 'Software\Microsoft\Office\PowerPoint\Addins\{#AddInKey}', 'FriendlyName', 'SlideBridge');
  RegWriteDWordValue(HKCU, 'Software\Microsoft\Office\PowerPoint\Addins\{#AddInKey}', 'LoadBehavior', 3);
  RegWriteStringValue(HKCU, 'Software\Microsoft\Office\PowerPoint\Addins\{#AddInKey}', 'Manifest', ManifestUrl + '|vstolocal');
  RegWriteStringValue(HKCU, 'Software\Kingsoft\Office\WPP\AddinsWL', '{#AddInKey}', '');
  RegWriteDWordValue(HKCU, 'Software\Microsoft\Office\16.0\PowerPoint\Resiliency\DoNotDisableAddinList', '{#AddInKey}', 1);

  if PublicKey <> '' then
  begin
    RegWriteStringValue(HKCU, InclusionKey, 'Url', ManifestUrl);
    RegWriteStringValue(HKCU, InclusionKey, 'PublicKey', PublicKey);
  end;
end;

procedure ResetCurrentVstoMetadata;
var
  ManifestUrl: string;
  ManifestPath: string;
  MetaRoot: string;
  SolutionId: string;
begin
  ManifestPath := ExpandConstant('{app}\{#VstoManifest}');
  ManifestUrl := FileUrl(ManifestPath);
  MetaRoot := 'Software\Microsoft\VSTO\SolutionMetadata';

  if RegQueryStringValue(HKCU, MetaRoot, ManifestUrl, SolutionId) then
  begin
    RegDeleteKeyIncludingSubkeys(HKCU, MetaRoot + '\' + SolutionId);
  end;

  RegDeleteValue(HKCU, MetaRoot, ManifestUrl);
end;

procedure ExpandDeployFiles(Dir: string);
var
  FindRec: TFindRec;
  SourcePath: string;
  DestPath: string;
begin
  if FindFirst(Dir + '\*', FindRec) then
  begin
    try
      repeat
        if (FindRec.Name <> '.') and (FindRec.Name <> '..') then
        begin
          SourcePath := Dir + '\' + FindRec.Name;
          if (FindRec.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0 then
          begin
            ExpandDeployFiles(SourcePath);
          end
          else if CompareText(ExtractFileExt(SourcePath), '.deploy') = 0 then
          begin
            DestPath := Copy(SourcePath, 1, Length(SourcePath) - Length('.deploy'));
            CopyFile(SourcePath, DestPath, False);
          end;
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;
end;

procedure CleanupOldApplicationFiles;
var
  FindRec: TFindRec;
  ApplicationFilesDir: string;
  CurrentVersionDir: string;
  CandidatePath: string;
  VersionPart: string;
begin
  ApplicationFilesDir := ExpandConstant('{app}\Application Files');
  VersionPart := '{#AppVersion}';
  StringChangeEx(VersionPart, '.', '_', True);
  CurrentVersionDir := 'SlideSCICompat_' + VersionPart;

  if FindFirst(ApplicationFilesDir + '\SlideSCICompat_*', FindRec) then
  begin
    try
      repeat
        if ((FindRec.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0) and
           (CompareText(FindRec.Name, CurrentVersionDir) <> 0) then
        begin
          CandidatePath := ApplicationFilesDir + '\' + FindRec.Name;
          DelTree(CandidatePath, True, True, True);
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;
end;

procedure CopyRuntimeFilesToAppRoot;
var
  FindRec: TFindRec;
  ApplicationFilesDir: string;
  CurrentVersionDir: string;
  SourceDir: string;
  SourcePath: string;
  DestPath: string;
  VersionPart: string;
  Ext: string;
begin
  ApplicationFilesDir := ExpandConstant('{app}\Application Files');
  VersionPart := '{#AppVersion}';
  StringChangeEx(VersionPart, '.', '_', True);
  CurrentVersionDir := 'SlideSCICompat_' + VersionPart;
  SourceDir := ApplicationFilesDir + '\' + CurrentVersionDir;

  if FindFirst(SourceDir + '\*', FindRec) then
  begin
    try
      repeat
        if ((FindRec.Attributes and FILE_ATTRIBUTE_DIRECTORY) = 0) then
        begin
          Ext := ExtractFileExt(FindRec.Name);
          if
            (CompareText(Ext, '.dll') = 0) or
            (CompareText(Ext, '.config') = 0) or
            (CompareText(Ext, '.manifest') = 0)
          then
          begin
            SourcePath := SourceDir + '\' + FindRec.Name;
            DestPath := ExpandConstant('{app}') + '\' + FindRec.Name;
            CopyFile(SourcePath, DestPath, False);
          end;
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;
end;

procedure CopyLatexConverterFilesToAppRoot;
var
  FindRec: TFindRec;
  ApplicationFilesDir: string;
  CurrentVersionDir: string;
  SourceDir: string;
  DestDir: string;
  SourcePath: string;
  DestPath: string;
  VersionPart: string;
begin
  ApplicationFilesDir := ExpandConstant('{app}\Application Files');
  VersionPart := '{#AppVersion}';
  StringChangeEx(VersionPart, '.', '_', True);
  CurrentVersionDir := 'SlideSCICompat_' + VersionPart;
  SourceDir := ApplicationFilesDir + '\' + CurrentVersionDir + '\latex-converter';
  DestDir := ExpandConstant('{app}\latex-converter');

  if not DirExists(SourceDir) then
  begin
    exit;
  end;

  ForceDirectories(DestDir);

  if FindFirst(SourceDir + '\*', FindRec) then
  begin
    try
      repeat
        if ((FindRec.Attributes and FILE_ATTRIBUTE_DIRECTORY) = 0) and
           (CompareText(ExtractFileExt(FindRec.Name), '.deploy') <> 0) then
        begin
          SourcePath := SourceDir + '\' + FindRec.Name;
          DestPath := DestDir + '\' + FindRec.Name;
          CopyFile(SourcePath, DestPath, False);
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    ExpandDeployFiles(ExpandConstant('{app}'));
    CleanupOldApplicationFiles;
    CopyRuntimeFilesToAppRoot;
    CopyLatexConverterFilesToAppRoot;
    RegisterVstoTrust;
  end;

  if CurStep = ssDone then
  begin
    RegisterVstoTrust;
  end;
end;

// Runtime DLLs are copied to the app root so VSTO can resolve the entry assembly.
// RegisterVstoTrust enforces Manifest, trust metadata and LoadBehavior.
