package main

import (
	"errors"
	"log"
	"net"
	"os"
	"os/exec"
	"strings"
	"sync"
	"time"
)

// app is the one running launcher: the presentation, the settings window and the tray icon,
// together. It lives while either is open and ends when both are gone (a person closed the
// presentation and there is no window), so the tray icon always means "the launcher is running".
type app struct {
	lock net.Listener

	mu       sync.Mutex
	pres     *presenter
	presDone chan struct{}
	busy     int // a presentation switch is in progress: never idle in between

	ui         *localUI
	window     *exec.Cmd
	windowLive bool
	lastPing   time.Time

	tray     *trayMenu // nil until the tray is up, or for good on a desktop without one
	quit     chan struct{}
	quitOnce sync.Once
}

func newApp(lock net.Listener) *app {
	a := &app{lock: lock, quit: make(chan struct{})}
	go a.serveControl()
	go func() {
		// The window can outlive the browser process we started (Chrome may hand it to one already
		// running), so a silent window also counts as closed.
		for range time.Tick(5 * time.Second) {
			a.exitIfIdle()
		}
	}()
	return a
}

// run blocks until the launcher ends.
func (a *app) run() {
	if trayAvailable() {
		runTray(a)
	} else {
		log.Printf("no tray on this desktop: running without the icon")
		<-a.quit
	}
	a.closeWindow()
}

func (a *app) serveControl() {
	for {
		conn, err := a.lock.Accept()
		if err != nil {
			return
		}
		command := strings.TrimSpace(readCommand(conn))
		conn.Close()
		log.Printf("command %q", command)
		switch command {
		case "show":
			if err := a.showSettings(); err != nil {
				log.Printf("show: %v", err)
			}
		case "start":
			if err := a.startPresentation(); err != nil {
				log.Printf("start: %v", err)
			}
		case "quit":
			a.exit()
		}
	}
}

func (a *app) presenting() bool {
	a.mu.Lock()
	defer a.mu.Unlock()
	return a.pres != nil
}

// startPresentation opens the saved address in kiosk mode, closing the current presentation first.
func (a *app) startPresentation() error {
	cfg := loadConfig()
	address, ok := normalizeURL(cfg.URL)
	if !ok {
		return errors.New("no valid Power Slider URL configured")
	}
	a.mu.Lock()
	a.busy++
	a.mu.Unlock()
	defer func() {
		a.mu.Lock()
		a.busy--
		a.mu.Unlock()
	}()

	a.stopPresentation()
	p := &presenter{browser: cfg.Browser, address: presentationURL(address)}
	done := make(chan struct{})
	a.mu.Lock()
	a.pres, a.presDone = p, done
	a.mu.Unlock()
	a.refreshTray()
	go func() {
		if err := p.watch(); err != nil {
			log.Printf("presentation: %v", err)
		}
		a.mu.Lock()
		if a.pres == p {
			a.pres = nil
		}
		a.mu.Unlock()
		close(done)
		a.refreshTray()
		a.exitIfIdle()
	}()
	return nil
}

// stopPresentation closes the browser and waits for the watchdog to finish.
func (a *app) stopPresentation() {
	a.mu.Lock()
	p, done := a.pres, a.presDone
	a.mu.Unlock()
	if p == nil {
		return
	}
	p.stop()
	<-done
}

func (a *app) restartPresentation() {
	a.mu.Lock()
	p := a.pres
	a.mu.Unlock()
	if p != nil {
		p.restart()
	}
}

// showSettings opens the settings window, or brings the open one to the front.
func (a *app) showSettings() error {
	a.mu.Lock()
	defer a.mu.Unlock()
	if a.windowOpenLocked() {
		bringToFront("Power Slider Launcher")
		return nil
	}
	if a.ui == nil {
		ui, err := newLocalUI()
		if err != nil {
			return err
		}
		a.registerSettings(ui)
		a.ui = ui
	}
	cmd, err := a.ui.open("index.html", 560, 640)
	if err != nil {
		return err
	}
	a.window, a.windowLive, a.lastPing = cmd, true, time.Now()
	go func() {
		_ = cmd.Wait()
		a.mu.Lock()
		if a.window == cmd {
			a.windowLive = false
		}
		a.mu.Unlock()
		a.exitIfIdle()
	}()
	return nil
}

func (a *app) windowOpenLocked() bool {
	return a.window != nil && (a.windowLive || time.Since(a.lastPing) < 15*time.Second)
}

func (a *app) ping() {
	a.mu.Lock()
	a.lastPing = time.Now()
	a.mu.Unlock()
}

func (a *app) closeWindow() {
	a.mu.Lock()
	w := a.window
	a.window, a.windowLive = nil, false
	a.mu.Unlock()
	if w != nil && w.Process != nil {
		_ = w.Process.Kill()
	}
}

// exitIfIdle ends the launcher once nothing is presenting and no window is open.
func (a *app) exitIfIdle() {
	a.mu.Lock()
	idle := a.pres == nil && a.busy == 0 && !a.windowOpenLocked()
	a.mu.Unlock()
	if idle {
		log.Printf("nothing presenting and no window: exiting")
		a.exit()
	}
}

func (a *app) exit() {
	a.quitOnce.Do(func() {
		go func() {
			a.stopPresentation()
			close(a.quit)
			quitTray()
			// Whatever is still blocking (a hung tray) must not keep a dead launcher around.
			time.Sleep(5 * time.Second)
			os.Exit(0)
		}()
	})
}
