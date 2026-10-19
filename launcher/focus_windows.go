package main

import (
	"syscall"
	"unsafe"
)

// bringToFront restores and raises the top-level window with this title (the settings window,
// whose title is its page's <title>).
func bringToFront(title string) {
	user32 := syscall.NewLazyDLL("user32.dll")
	t, err := syscall.UTF16PtrFromString(title)
	if err != nil {
		return
	}
	hwnd, _, _ := user32.NewProc("FindWindowW").Call(0, uintptr(unsafe.Pointer(t)))
	if hwnd == 0 {
		return
	}
	const swRestore = 9
	user32.NewProc("ShowWindow").Call(hwnd, swRestore)
	user32.NewProc("SetForegroundWindow").Call(hwnd)
}
