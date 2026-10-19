; Windows installer. build.ps1 passes the version: ISCC /DAppVersion=x.y.z inno_setup.iss
#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif

[Setup]
AppName=Power Slider Launcher
AppVersion={#AppVersion}
VersionInfoVersion={#AppVersion}
AppPublisher=Atomo
AppPublisherURL=https://www.powerslider.app
; Mesmo AppId de sempre: o instalador novo atualiza o antigo no lugar.
AppId=12951e62-75ce-4dc1-9ef8-fcb629373ed4
WizardStyle=modern
; A pasta e o executável mantêm o nome antigo: o atalho da pasta Inicializar de cada tela aponta
; para "{app}\PB Slider Launcher.exe launch", e é ele que abre a apresentação no boot.
DefaultDirName={commonpf}\PB Slider Launcher
OutputDir=dist\release
OutputBaseFilename=Power_Slider_Launcher_Setup
SetupIconFile=logo.ico
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\logo.ico
DefaultGroupName=Power Slider Launcher
; Sem isso, uma atualização reaproveitaria o grupo antigo do menu Iniciar.
UsePreviousGroup=no
; A atualização automática roda este instalador com o launcher aberto: ele fecha o processo.
CloseApplications=force
RestartApplications=no

[InstallDelete]
; Atalhos com o nome antigo, de instalações anteriores. O da pasta Inicializar não entra aqui.
Type: filesandordirs; Name: "{commonprograms}\PB Slider Launcher"
Type: files; Name: "{commondesktop}\PB Slider Launcher.lnk"

[Files]
Source: "dist\build\windows-amd64\PB Slider Launcher.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "logo.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{commondesktop}\Power Slider Launcher"; Filename: "{app}\PB Slider Launcher.exe"; IconFilename: "{app}\logo.ico"
Name: "{group}\Power Slider Launcher"; Filename: "{app}\PB Slider Launcher.exe"; IconFilename: "{app}\logo.ico"
Name: "{group}\Uninstall Power Slider Launcher"; Filename: "{uninstallexe}"; IconFilename: "{app}\logo.ico"

[Run]
; Instalação manual: abre a configuração no fim.
Filename: "{app}\PB Slider Launcher.exe"; Description: "Open Power Slider Launcher"; Flags: nowait postinstall skipifsilent runasoriginaluser
; Atualização automática (silenciosa): volta ao que estava aberto, /RELAUNCH=launch ou settings.
Filename: "{app}\PB Slider Launcher.exe"; Parameters: "launch"; Flags: nowait runasoriginaluser; Check: RelaunchIs('launch')
Filename: "{app}\PB Slider Launcher.exe"; Flags: nowait runasoriginaluser; Check: RelaunchIs('settings')

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
Type: filesandordirs; Name: "{localappdata}\PB Slider Launcher"

[Registry]
; Remove entradas antigas de início automático
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueName: "PB Slider Launcher"; Flags: deletevalue uninsdeletevalue
Root: HKLM; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueName: "PB Slider Launcher"; Flags: deletevalue uninsdeletevalue

[UninstallRun]
Filename: "{cmd}"; Parameters: "/C if exist ""{userstartup}\PB Slider Launcher.lnk"" del ""{userstartup}\PB Slider Launcher.lnk"""; Flags: runhidden; RunOnceId: "DelUserStartup"
Filename: "{cmd}"; Parameters: "/C if exist ""{commonstartup}\PB Slider Launcher.lnk"" del ""{commonstartup}\PB Slider Launcher.lnk"""; Flags: runhidden; RunOnceId: "DelCommonStartup"

[Code]
function RelaunchIs(Mode: String): Boolean;
begin
  Result := WizardSilent and (CompareText(ExpandConstant('{param:RELAUNCH|}'), Mode) = 0);
end;
