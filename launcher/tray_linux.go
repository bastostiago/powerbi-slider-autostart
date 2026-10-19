package main

import (
	_ "embed"
	"os"

	"github.com/godbus/dbus/v5"
)

//go:embed ui/icon.png
var trayIcon []byte

// systemLanguage reads the locale the desktop session runs in, e.g. "pt_BR.UTF-8".
func systemLanguage() string {
	for _, name := range []string{"LC_ALL", "LC_MESSAGES", "LANG"} {
		if v := os.Getenv(name); v != "" && v != "C" && v != "POSIX" {
			return v
		}
	}
	return "en"
}

// trayAvailable reports whether the session has a StatusNotifierItem host to show the icon. Without
// one, the tray library also misbehaves: with no session bus at all it panics while shutting down.
func trayAvailable() bool {
	conn, err := dbus.ConnectSessionBus()
	if err != nil {
		return false
	}
	defer conn.Close()
	var owned bool
	err = conn.BusObject().Call("org.freedesktop.DBus.NameHasOwner", 0, "org.kde.StatusNotifierWatcher").Store(&owned)
	return err == nil && owned
}
