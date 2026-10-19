[Setup]
AppName=Power Slider Launcher
AppVersion=2.0.0
; Mesmo AppId de sempre: o instalador novo atualiza o antigo no lugar.
AppId=12951e62-75ce-4dc1-9ef8-fcb629373ed4
WizardStyle=modern
; A pasta e o executável mantêm o nome antigo: o atalho da pasta Inicializar de cada tela aponta
; para "{app}\PB Slider Launcher.exe", e é ele que abre a apresentação no boot.
DefaultDirName={commonpf}\PB Slider Launcher
OutputBaseFilename=Power_Slider_Launcher_Installer
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\logo.ico
DefaultGroupName=Power Slider Launcher
; Sem isso, uma atualização reaproveitaria o grupo antigo do menu Iniciar.
UsePreviousGroup=no

[InstallDelete]
; Atalhos com o nome antigo, de instalações anteriores. O da pasta Inicializar não entra aqui.
Type: filesandordirs; Name: "{commonprograms}\PB Slider Launcher"
Type: files; Name: "{commondesktop}\PB Slider Launcher.lnk"

[Files]
Source: "dist\PB Slider Launcher.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "logo.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Atalho na área de trabalho
Name: "{commondesktop}\Power Slider Launcher"; Filename: "{app}\PB Slider Launcher.exe"; IconFilename: "{app}\logo.ico"

; Atalho no menu Iniciar (mas não no Startup Folder)
Name: "{group}\Power Slider Launcher"; Filename: "{app}\PB Slider Launcher.exe"; IconFilename: "{app}\logo.ico"

; Atalho para desinstalar
Name: "{group}\Uninstall Power Slider Launcher"; Filename: "{uninstallexe}"; IconFilename: "{app}\logo.ico"

[Run]
; Executa o aplicativo após a instalação (opcional)
Filename: "{app}\PB Slider Launcher.exe"; Description: "Execute Power Slider Launcher"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Remove a pasta de instalação
Type: filesandordirs; Name: "{app}"
Type: filesandordirs; Name: "{localappdata}\PB Slider Launcher"

[Registry]
; Remove any potential auto-start entries
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueName: "PB Slider Launcher"; Flags: deletevalue uninsdeletevalue
Root: HKLM; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueName: "PB Slider Launcher"; Flags: deletevalue uninsdeletevalue

[UninstallRun]
Filename: "{cmd}"; Parameters: "/C if exist ""{userstartup}\PB Slider Launcher.lnk"" del ""{userstartup}\PB Slider Launcher.lnk"""; Flags: runhidden
Filename: "{cmd}"; Parameters: "/C if exist ""{commonstartup}\PB Slider Launcher.lnk"" del ""{commonstartup}\PB Slider Launcher.lnk"""; Flags: runhidden
