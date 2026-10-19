package main

import (
	"archive/tar"
	"compress/gzip"
	"errors"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
)

func dataDir() string {
	if dir := os.Getenv("PSL_DATA_DIR"); dir != "" {
		return dir // tests and trials, away from a real screen's setup
	}
	if dir := os.Getenv("XDG_CONFIG_HOME"); dir != "" {
		return filepath.Join(dir, "power-slider-launcher")
	}
	home, _ := os.UserHomeDir()
	return filepath.Join(home, ".config", "power-slider-launcher")
}

func browserCandidates(name string) []string {
	switch name {
	case "chrome":
		return []string{"google-chrome", "google-chrome-stable"}
	case "edge":
		return []string{"microsoft-edge", "microsoft-edge-stable"}
	case "chromium":
		return []string{"chromium", "chromium-browser"}
	}
	return nil
}

// An XDG autostart entry starts the launcher when the desktop session opens, which is when a
// browser can show a window at all.
func startupShortcut() string {
	dir := os.Getenv("XDG_CONFIG_HOME")
	if dir == "" {
		home, _ := os.UserHomeDir()
		dir = filepath.Join(home, ".config")
	}
	return filepath.Join(dir, "autostart", "power-slider-launcher.desktop")
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
	if err := os.MkdirAll(filepath.Dir(startupShortcut()), 0o755); err != nil {
		return err
	}
	entry := fmt.Sprintf("[Desktop Entry]\nType=Application\nName=Power Slider Launcher\nExec=%q launch\nX-GNOME-Autostart-enabled=true\n", exe)
	return os.WriteFile(startupShortcut(), []byte(entry), 0o644)
}

// A launcher under /usr came from the .deb and is updated by one; anywhere else it was unpacked
// from the tarball and replaces itself.
func installedByPackage() bool {
	exe, err := os.Executable()
	return err == nil && strings.HasPrefix(exe, "/usr/")
}

func updateAssetName() string {
	if installedByPackage() {
		return fmt.Sprintf("power-slider-launcher_%s.deb", runtime.GOARCH)
	}
	return fmt.Sprintf("power-slider-launcher-linux-%s.tar.gz", runtime.GOARCH)
}

func installUpdate(path, relaunch string) error {
	exe, err := os.Executable()
	if err != nil {
		return err
	}
	if strings.HasSuffix(path, ".deb") {
		// pkexec asks for the administrator password in a desktop dialog.
		if out, err := exec.Command("pkexec", "apt-get", "install", "-y", path).CombinedOutput(); err != nil {
			return fmt.Errorf("apt-get: %v: %s", err, out)
		}
	} else if err := replaceFromTarball(path, exe); err != nil {
		return err
	}
	if relaunch == "" {
		return nil
	}
	arg := "launch"
	if relaunch == "settings" {
		arg = ""
	}
	cmd := exec.Command(exe)
	if arg != "" {
		cmd.Args = append(cmd.Args, arg)
	}
	return cmd.Start()
}

func replaceFromTarball(tarball, exe string) error {
	f, err := os.Open(tarball)
	if err != nil {
		return err
	}
	defer f.Close()
	gz, err := gzip.NewReader(f)
	if err != nil {
		return err
	}
	tr := tar.NewReader(gz)
	for {
		h, err := tr.Next()
		if err == io.EOF {
			return errors.New("tarball has no power-slider-launcher binary")
		}
		if err != nil {
			return err
		}
		if filepath.Base(h.Name) != "power-slider-launcher" || h.Typeflag != tar.TypeReg {
			continue
		}
		tmp := exe + ".new"
		out, err := os.OpenFile(tmp, os.O_CREATE|os.O_TRUNC|os.O_WRONLY, 0o755)
		if err != nil {
			return err
		}
		if _, err := io.Copy(out, tr); err != nil {
			out.Close()
			return err
		}
		if err := out.Close(); err != nil {
			return err
		}
		// A running binary can be replaced by rename on Linux; the old inode lives until exit.
		return os.Rename(tmp, exe)
	}
}

func hideWindow(*exec.Cmd) {}
