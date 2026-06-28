#define MyAppName "OfficeFileGuard"
#define MyAppVersion "2.0.0"
#define MyAppPublisher "BlueCrocodile"
#define MyAppExeName "CrocoOfficeFileReserveToggle.Cli.exe"
#define SourceBase "{#SourcePath}\..\.."

[Setup]
;nuove aggiunte
VersionInfoVersion={#MyAppVersion}
;VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName} Installer
VersionInfoProductName={#MyAppName}
;VersionInfoCopyright=Copyright © 2026 BlueCrocodile
;AppId={{B8E5A5F4-72C3-4D5F-9C8E-1A2B3C4D5E6F}
;DisableDirPage=auto
;UsePreviousAppDir=yes
; fine nuove aggiunte

MinVersion=10.0
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir={#SourceBase}\installer\output
OutputBaseFilename=OfficeFileGuard_Setup
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=admin
UninstallFilesDir={app}
UninstallDisplayName=OfficeFileGuard
UninstallDisplayIcon={app}\lucchetto_tr.ico
LicenseFile=license.txt
; riavvio di tutte le applicazioni coinvolte nell'installazione (incluso windows explorer)
;CloseApplications=force
;RestartApplications=yes
; Riavvia Windows al termine dell'installazione
;AlwaysRestart=yes

[Files]
; CLI
Source: "{#SourceBase}\src\CrocoOfficeFileReserveToggle.Cli\bin\x64\Release\CrocoOfficeFileReserveToggle.Cli.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceBase}\src\CrocoOfficeFileReserveToggle.Cli\bin\x64\Release\CrocoOfficeFileGuard.Core.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceBase}\src\CrocoOfficeFileReserveToggle.Cli\bin\x64\Release\DocumentFormat.OpenXml.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceBase}\src\CrocoOfficeFileReserveToggle.Cli\bin\x64\Release\DocumentFormat.OpenXml.Framework.dll"; DestDir: "{app}"; Flags: ignoreversion

; Overlay icon shell extension
Source: "{#SourceBase}\CrocoIconOverlayShellExtension\bin\x64\Release\CrocoIconOverlayShellExtension.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceBase}\CrocoIconOverlayShellExtension\properties\lucchetto_tr.ico"; DestDir: "{app}"; Flags: ignoreversion

; Outlook add-in
Source: "{#SourceBase}\src\OutlookFileGuard_addin\bin\x64\Release\OutlookFileGuard_addin.dll"; DestDir: "{app}\addin"; Flags: ignoreversion
Source: "{#SourceBase}\src\OutlookFileGuard_addin\bin\x64\Release\OutlookFileGuard_addin.dll.manifest"; DestDir: "{app}\addin"; Flags: ignoreversion
Source: "{#SourceBase}\src\OutlookFileGuard_addin\bin\x64\Release\OutlookFileGuard_addin.vsto"; DestDir: "{app}\addin"; Flags: ignoreversion
Source: "{#SourceBase}\src\OutlookFileGuard_addin\bin\x64\Release\CrocoOfficeFileGuard.Core.dll"; DestDir: "{app}\addin"; Flags: ignoreversion
Source: "{#SourceBase}\src\OutlookFileGuard_addin\bin\x64\Release\DocumentFormat.OpenXml.dll"; DestDir: "{app}\addin"; Flags: ignoreversion
Source: "{#SourceBase}\src\OutlookFileGuard_addin\bin\x64\Release\DocumentFormat.OpenXml.Framework.dll"; DestDir: "{app}\addin"; Flags: ignoreversion
Source: "{#SourceBase}\src\OutlookFileGuard_addin\bin\x64\Release\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"; DestDir: "{app}\addin"; Flags: ignoreversion
Source: "{#SourceBase}\src\OutlookFileGuard_addin\bin\x64\Release\Microsoft.Office.Tools.Outlook.v4.0.Utilities.dll"; DestDir: "{app}\addin"; Flags: ignoreversion

; QuickStart e License
Source: "license.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "QuickStart.txt"; DestDir: "{app}"; Flags: ignoreversion

[Registry]

; ==================== CONTEXT MENUS ====================

; Menu contestuale per .docx
Root: HKLM; Subkey: "SOFTWARE\Classes\SystemFileAssociations\.docx\shell\ToggleReserved"; ValueType: string; ValueName: ""; ValueData: "Toggle Reserved"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Classes\SystemFileAssociations\.docx\shell\ToggleReserved\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Classes\SystemFileAssociations\.docx\shell\ToggleReserved"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\lucchetto_tr.ico"; Flags: uninsdeletevalue
Root: HKLM; Subkey: "SOFTWARE\Classes\SystemFileAssociations\.docx\shell\ToggleReserved"; ValueType: string; ValueName: "Position"; ValueData: "Top"; Flags: uninsdeletevalue

; Menu contestuale per .xlsx
Root: HKLM; Subkey: "SOFTWARE\Classes\SystemFileAssociations\.xlsx\shell\ToggleReserved"; ValueType: string; ValueName: ""; ValueData: "Toggle Reserved"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Classes\SystemFileAssociations\.xlsx\shell\ToggleReserved\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Classes\SystemFileAssociations\.xlsx\shell\ToggleReserved"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\lucchetto_tr.ico"; Flags: uninsdeletevalue
Root: HKLM; Subkey: "SOFTWARE\Classes\SystemFileAssociations\.xlsx\shell\ToggleReserved"; ValueType: string; ValueName: "Position"; ValueData: "Top"; Flags: uninsdeletevalue

; Menu contestuale per .pptx
Root: HKLM; Subkey: "SOFTWARE\Classes\SystemFileAssociations\.pptx\shell\ToggleReserved"; ValueType: string; ValueName: ""; ValueData: "Toggle Reserved"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Classes\SystemFileAssociations\.pptx\shell\ToggleReserved\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Classes\SystemFileAssociations\.pptx\shell\ToggleReserved"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\lucchetto_tr.ico"; Flags: uninsdeletevalue
Root: HKLM; Subkey: "SOFTWARE\Classes\SystemFileAssociations\.pptx\shell\ToggleReserved"; ValueType: string; ValueName: "Position"; ValueData: "Top"; Flags: uninsdeletevalue


; ==================== SHELL ICON OVERLAY ====================

; Overlay Icon Handler (con spazio iniziale per priorità alta)
;Root: HKLM; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\ CrocoOfficeOverlayHandler"; ValueType: string; ValueName: ""; ValueData: "{{098AA80F-2F55-4AA8-B95F-8AF77120E479}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\     !CrocoOffice"; ValueType: string; ValueName: ""; ValueData: "{{098AA80F-2F55-4AA8-B95F-8AF77120E479}"; Flags: uninsdeletekey

; ==================== CLSID ====================

Root: HKLM; Subkey: "Software\Classes\CLSID\{{098AA80F-2F55-4AA8-B95F-8AF77120E479}"; ValueType: string; ValueName: ""; ValueData: "CrocoOfficeOverlayHandler"; Flags: uninsdeletekey

Root: HKLM; Subkey: "Software\Classes\CLSID\{{098AA80F-2F55-4AA8-B95F-8AF77120E479}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Classes\CLSID\{{098AA80F-2F55-4AA8-B95F-8AF77120E479}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "CrocoIconOverlayShellExtension, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"; Flags: uninsdeletevalue
Root: HKLM; Subkey: "Software\Classes\CLSID\{{098AA80F-2F55-4AA8-B95F-8AF77120E479}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "CrocoIconOverlayShellExtension.CrocoOfficeOverlayHandler"; Flags: uninsdeletevalue
Root: HKLM; Subkey: "Software\Classes\CLSID\{{098AA80F-2F55-4AA8-B95F-8AF77120E479}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{app}\CrocoIconOverlayShellExtension.dll"; Flags: uninsdeletevalue
Root: HKLM; Subkey: "Software\Classes\CLSID\{{098AA80F-2F55-4AA8-B95F-8AF77120E479}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Flags: uninsdeletevalue
Root: HKLM; Subkey: "Software\Classes\CLSID\{{098AA80F-2F55-4AA8-B95F-8AF77120E479}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Flags: uninsdeletevalue

; Version-specific keys
Root: HKLM; Subkey: "Software\Classes\CLSID\{{098AA80F-2F55-4AA8-B95F-8AF77120E479}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "CrocoIconOverlayShellExtension, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"; Flags: uninsdeletevalue
Root: HKLM; Subkey: "Software\Classes\CLSID\{{098AA80F-2F55-4AA8-B95F-8AF77120E479}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "Class"; ValueData: "CrocoIconOverlayShellExtension.CrocoOfficeOverlayHandler"; Flags: uninsdeletevalue
Root: HKLM; Subkey: "Software\Classes\CLSID\{{098AA80F-2F55-4AA8-B95F-8AF77120E479}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{app}\CrocoIconOverlayShellExtension.dll"; Flags: uninsdeletevalue
Root: HKLM; Subkey: "Software\Classes\CLSID\{{098AA80F-2F55-4AA8-B95F-8AF77120E479}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Flags: uninsdeletevalue


; ==================== OUTLOOK ADD-IN ====================

Root: HKLM; Subkey: "Software\Microsoft\Office\Outlook\Addins\OutlookFileGuard_addin"; ValueType: string; ValueName: "Description"; ValueData: "OutlookFileGuard_addin"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\Outlook\Addins\OutlookFileGuard_addin"; ValueType: string; ValueName: "FriendlyName"; ValueData: "OutlookFileGuard_addin"; Flags: uninsdeletevalue
Root: HKLM; Subkey: "Software\Microsoft\Office\Outlook\Addins\OutlookFileGuard_addin"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletevalue
Root: HKLM; Subkey: "Software\Microsoft\Office\Outlook\Addins\OutlookFileGuard_addin"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\addin\OutlookFileGuard_addin.vsto|vstolocal"; Flags: uninsdeletevalue

[Run]
; Riavvia Explorer dopo l'installazione per caricare la shell extension
;Filename: "{sys}\taskkill.exe"; Parameters: "/f /im explorer.exe"; Flags: waituntilterminated
;Filename: "{win}\explorer.exe"; Flags: nowait runasoriginaluser

;Filename: "{cmd}"; Parameters: "/c taskkill /f /im explorer.exe & ping 127.0.0.1 -n 3 >nul & start explorer.exe"; Flags: runhidden

; Mostra Quick Start al termine dell'installazione
;Filename: "{app}\QuickStart.txt"; Description: "Show QuickStart"; Flags: postinstall shellexec skipifsilent

[UninstallRun]
; Riavvia Explorer dopo la disinstallazione
Filename: "{cmd}"; Parameters: "/c taskkill /f /im explorer.exe & ping 127.0.0.1 -n 3 >nul & start explorer.exe"; Flags: runhidden; RunOnceId: "RestartExplorerUninstall"
;Filename: "{sys}\taskkill.exe"; Parameters: "/f /im explorer.exe"; Flags: waituntilterminated; RunOnceId: "RestartExplorerUninstall"
;3 sec delay
;Filename: "{sys}\ping.exe"; Parameters: "-n 3 127.0.0.1"; Flags: waituntilterminated runhidden; RunOnceId: "DelayUninstall"  
;Filename: "{win}\explorer.exe"; Flags: nowait; RunOnceId: "StartExplorerUninstall"

[Code]
function IsDotNet48Installed(): Boolean;
var
  Release: Cardinal;
begin
  Result := False;
  if RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Release) then
    Result := Release >= 528040; // 528040 = .NET 4.8
end;

function InitializeSetup(): Boolean;
var
  Msg: String;
  MissingComponents: Boolean;
begin
  Result := True;
  MissingComponents := False;
  Msg := 'The following required components are missing:' + #13#10 + #13#10;

  if not IsDotNet48Installed() then
  begin
    MissingComponents := True;
    Msg := Msg + '- .NET Framework 4.8' + #13#10;
    Msg := Msg + '  Download from: https://dotnet.microsoft.com/download/dotnet-framework/net48' + #13#10 + #13#10;
  end;

  if MissingComponents then
  begin
    Msg := Msg + 'Please install the missing components and then run this installer again.';
    MsgBox(Msg, mbError, MB_OK);
    Result := False; // blocca l'installazione
  end;
end;

[Code]
function NeedRestart(): Boolean;
begin
  Result := True;
end;

// MessageBox che chiede il restart
(*
procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
begin
  if CurStep = ssPostInstall then
  begin
    if MsgBox('Setup has completed the installation! You must restart your computer for the changes to take effect. Do you want to restart now?', mbConfirmation, MB_YESNO) = IDYES then
    begin
      Exec('shutdown.exe', '/r /t 0', '', SW_HIDE, ewNoWait, ResultCode);
    end;
  end;
end;
*)