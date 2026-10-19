package main

import (
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"syscall"
)

// The old name on purpose: config.ini and the browser profiles (with each screen's Microsoft
// session) already live there.
func dataDir() string {
	if dir := os.Getenv("PSL_DATA_DIR"); dir != "" {
		return dir // tests and trials, away from a real screen's setup
	}
	return filepath.Join(os.Getenv("LOCALAPPDATA"), "PB Slider Launcher")
}

func browserCandidates(name string) []string {
	pf, pf86, local := os.Getenv("ProgramFiles"), os.Getenv("ProgramFiles(x86)"), os.Getenv("LOCALAPPDATA")
	switch name {
	case "chrome":
		return []string{
			filepath.Join(pf, `Google\Chrome\Application\chrome.exe`),
			filepath.Join(pf86, `Google\Chrome\Application\chrome.exe`),
			filepath.Join(local, `Google\Chrome\Application\chrome.exe`),
		}
	case "edge":
		return []string{
			filepath.Join(pf86, `Microsoft\Edge\Application\msedge.exe`),
			filepath.Join(pf, `Microsoft\Edge\Application\msedge.exe`),
		}
	}
	return nil
}

// The shortcut name and its "launch" argument are the old launcher's: screens updated in place keep
// booting through the shortcut they already have.
func startupShortcut() string {
	return filepath.Join(os.Getenv("APPDATA"), `Microsoft\Windows\Start Menu\Programs\Startup`, "PB Slider Launcher.lnk")
}

func autostartEnabled() bool {
	_, err := os.Stat(startupShortcut())
	return err == nil
}

func setAutostart(enabled bool) error {
	if !enabled {
		err := os.Remove(startupShortcut())
		if os.IsNotExist(err) {
			return nil
		}
		return err
	}
	exe, err := os.Executable()
	if err != nil {
		return err
	}
	quote := func(s string) string { return "'" + strings.ReplaceAll(s, "'", "''") + "'" }
	script := strings.Join([]string{
		"$s = (New-Object -ComObject WScript.Shell).CreateShortcut(" + quote(startupShortcut()) + ")",
		"$s.TargetPath = " + quote(exe),
		"$s.Arguments = 'launch'",
		"$s.WorkingDirectory = " + quote(filepath.Dir(exe)),
		"$s.IconLocation = " + quote(exe),
		"$s.Save()",
	}, "; ")
	cmd := exec.Command("powershell", "-NoProfile", "-NonInteractive", "-Command", script)
	hideWindow(cmd)
	return cmd.Run()
}

func updateAssetName() string { return "Power_Slider_Launcher_Setup.exe" }

// installUpdate runs the new installer silently. It replaces the running executable, so the
// launcher exits right after starting it; /RELAUNCH tells the installer what to start again
// ("launch" for the presentation, "settings" for this window).
func installUpdate(path, relaunch string) error {
	args := []string{"/SILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/CLOSEAPPLICATIONS", "/FORCECLOSEAPPLICATIONS"}
	if relaunch != "" {
		args = append(args, "/RELAUNCH="+relaunch)
	}
	return exec.Command(path, args...).Start()
}

func hideWindow(cmd *exec.Cmd) {
	cmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true, CreationFlags: 0x08000000} // CREATE_NO_WINDOW
}
