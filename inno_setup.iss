[Setup]
AppName=PB Slider Launcher
AppVersion=1.0.0
WizardStyle=modern
DefaultDirName={commonpf}\PB Slider Launcher
OutputBaseFilename=PB_Slider_Launcher_Installer
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\logo.ico  
DefaultGroupName=PB Slider Launcher  

[Files]
Source: "dist\PB Slider Launcher.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "logo.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Atalho na área de trabalho
Name: "{commondesktop}\PB Slider Launcher"; Filename: "{app}\PB Slider Launcher.exe"; IconFilename: "{app}\logo.ico"

; Atalho no menu Iniciar (mas não no Startup Folder)
Name: "{group}\PB Slider Launcher"; Filename: "{app}\PB Slider Launcher.exe"; IconFilename: "{app}\logo.ico"

; Atalho para desinstalar
Name: "{group}\Uninstall PB Slider Launcher"; Filename: "{uninstallexe}"; IconFilename: "{app}\logo.ico"

[Run]
; Executa o aplicativo após a instalação (opcional)
Filename: "{app}\PB Slider Launcher.exe"; Description: "Execute PB Slider Launcher"; Flags: nowait postinstall skipifsilent

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
