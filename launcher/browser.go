package main

import (
	"errors"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
)

type browser struct {
	Name string // "chrome", "edge" or "chromium"
	Path string
}

// findBrowser resolves the configured browser, or the first one installed for "auto". Chrome
// comes first on Windows because that is what most screens set up with the old launcher used, and
// its profile holds their Microsoft session.
func findBrowser(preference string) (browser, error) {
	order := []string{"chrome", "edge", "chromium"}
	if preference != "" && preference != "auto" {
		order = []string{preference}
	}
	for _, name := range order {
		for _, candidate := range browserCandidates(name) {
			if path := resolve(candidate); path != "" {
				return browser{Name: name, Path: path}, nil
			}
		}
	}
	return browser{}, errors.New("no supported browser found (Google Chrome, Microsoft Edge or Chromium)")
}

func resolve(candidate string) string {
	if filepath.IsAbs(candidate) {
		if info, err := os.Stat(candidate); err == nil && !info.IsDir() {
			return candidate
		}
		return ""
	}
	path, err := exec.LookPath(candidate)
	if err != nil {
		return ""
	}
	return path
}

// profileDir is the browser's own user-data-dir: the Microsoft sign-in and the site's per-report
// settings live in it. The names match the old launcher's so an update keeps both.
func profileDir(b browser) string {
	switch b.Name {
	case "edge":
		return filepath.Join(dataDir(), "EdgeUserData")
	case "chromium":
		return filepath.Join(dataDir(), "ChromiumUserData")
	default:
		return filepath.Join(dataDir(), "ChromeUserData")
	}
}

// noBackground keeps the browser from living on after its last window closes (Edge's background
// mode and startup boost do, by default): the launcher learns that a person closed the presentation,
// or the settings window, from the process exiting. Chromium keeps only the last
// --disable-features, so every feature to disable goes in this one.
var noBackground = []string{
	"--disable-background-mode",
	"--disable-features=Translate,msEdgeStartupBoost,StartupBoost",
}

var commonArgs = append([]string{
	"--no-first-run",
	"--no-default-browser-check",
	// A kiosk that lost power would otherwise come back with a "restore pages?" bubble.
	"--disable-session-crashed-bubble",
	"--hide-crash-restore-bubble",
	"--noerrdialogs",
	"--autoplay-policy=no-user-gesture-required",
}, noBackground...)

// kioskCommand opens the presentation fullscreen, without the browser's toolbar.
func kioskCommand(b browser, address string) *exec.Cmd {
	args := append([]string{"--user-data-dir=" + profileDir(b)}, commonArgs...)
	if b.Name == "edge" {
		// Not Edge's --kiosk: both of its kiosk types run InPrivate, so the Microsoft sign-in was
		// gone after every restart. A fullscreen app window keeps the profile and shows no
		// browser UI either.
		return exec.Command(b.Path, append(args, "--start-fullscreen", "--app="+address)...)
	}
	return exec.Command(b.Path, append(args, "--kiosk", address)...)
}

// appWindowCommand shows a launcher page as a small app window with its own throwaway profile.
func appWindowCommand(b browser, address string, width, height int) *exec.Cmd {
	return exec.Command(b.Path,
		"--user-data-dir="+filepath.Join(dataDir(), "SettingsWindow"),
		"--app="+address,
		fmt.Sprintf("--window-size=%d,%d", width, height),
		"--no-first-run",
		"--no-default-browser-check",
		noBackground[0],
		noBackground[1],
	)
}
