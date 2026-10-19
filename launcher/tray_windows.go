package main

import (
	_ "embed"
	"syscall"
	"unsafe"
)

// Windows wants an .ico for the notification area.
//
//go:embed trayicon.ico
var trayIcon []byte

// The notification area is always there on Windows.
func trayAvailable() bool { return true }

// systemLanguage is the user's Windows display locale, e.g. "pt-BR".
func systemLanguage() string {
	buf := make([]uint16, 85) // LOCALE_NAME_MAX_LENGTH
	proc := syscall.NewLazyDLL("kernel32.dll").NewProc("GetUserDefaultLocaleName")
	if n, _, _ := proc.Call(uintptr(unsafe.Pointer(&buf[0])), uintptr(len(buf))); n == 0 {
		return "en"
	}
	return syscall.UTF16ToString(buf)
}
