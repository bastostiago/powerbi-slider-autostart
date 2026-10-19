// Power Slider Launcher: opens a Power Slider presentation fullscreen when the computer starts,
// and opens it again if the browser closes or crashes.
//
//	launcher          the settings window
//	launcher launch   the presentation, in kiosk mode, watched (what the boot shortcut runs)
package main

import (
	"bufio"
	"context"
	"fmt"
	"log"
	"net"
	"os"
	"os/exec"
	"path/filepath"
	"sync"
	"time"
)

// Set at build time from the VERSION file (see build.ps1).
var version = "dev"

func main() {
	_ = os.MkdirAll(dataDir(), 0o755)
	if f, err := os.OpenFile(filepath.Join(dataDir(), "launcher.log"), os.O_CREATE|os.O_APPEND|os.O_WRONLY, 0o644); err == nil {
		log.SetOutput(f)
		defer f.Close()
	}
	log.Printf("Power Slider Launcher %s, args %v", version, os.Args[1:])

	if len(os.Args) > 1 && os.Args[1] == "--version" {
		fmt.Println(version)
		return
	}
	if len(os.Args) > 1 && os.Args[1] == "launch" {
		if err := present(); err != nil {
			log.Printf("launch: %v", err)
			os.Exit(1)
		}
		return
	}
	if err := runSettings(); err != nil {
		log.Printf("settings: %v", err)
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}

// Watchdog limits: a browser that exits this many times inside the window was closed on purpose
// (or cannot start), and reopening it forever would fight the person at the keyboard.
const (
	maxQuickExits   = 3
	quickExitWindow = 2 * time.Minute
)

// presentLock is held by the one launcher that is presenting. A second "launch" (Start now clicked
// again, the shortcut run by hand) would hand its kiosk window to the running browser, which exits
// at once, and open a window per retry. A loopback port is the lock: the OS frees it when the
// process dies, however it dies. It doubles as a control channel: "quit" stops the presentation
// (the updater uses it before replacing the executable).
const presentLock = "127.0.0.1:47613"

func present() error {
	lock, err := net.Listen("tcp", presentLock)
	if err != nil {
		log.Printf("another launcher is already presenting: %v", err)
		return nil
	}
	defer lock.Close()

	cfg := loadConfig()
	address, ok := normalizeURL(cfg.URL)
	if !ok {
		return fmt.Errorf("no valid Power Slider URL configured (%q)", cfg.URL)
	}
	address = presentationURL(address)

	// At boot someone may be in front of the screen: offer a pending update before presenting. An
	// unanswered prompt times out and the presentation goes on.
	if offerUpdateAtBoot(func() { lock.Close() }) {
		return nil
	}

	p := &presenter{browser: cfg.Browser, address: address}
	go p.serveControl(lock)
	// The tray icon owns the main goroutine (Windows needs its message loop there); the watchdog
	// runs beside it and closes the tray when it stops.
	return runTray(p)
}

// presenter is the watchdog: it keeps the kiosk browser open until told to stop.
type presenter struct {
	browser string
	address string

	mu         sync.Mutex
	current    *exec.Cmd
	quitting   bool
	restarting bool
}

// stop closes the browser and ends the watchdog (the tray's "Stop presentation", the updater).
func (p *presenter) stop() {
	p.mu.Lock()
	defer p.mu.Unlock()
	p.quitting = true
	p.killCurrent()
}

// restart closes the browser and opens it again at once, without counting it as a close.
func (p *presenter) restart() {
	p.mu.Lock()
	defer p.mu.Unlock()
	p.restarting = true
	p.killCurrent()
}

func (p *presenter) killCurrent() {
	if p.current != nil && p.current.Process != nil {
		_ = p.current.Process.Kill()
	}
}

func (p *presenter) serveControl(lock net.Listener) {
	for {
		conn, err := lock.Accept()
		if err != nil {
			return
		}
		line, _ := bufio.NewReader(conn).ReadString('\n')
		if line == "quit\n" {
			log.Printf("asked to stop presenting")
			p.stop()
		}
		conn.Close()
	}
}

func (p *presenter) watch() error {
	var exits []time.Time
	delay := 2 * time.Second
	for {
		// Looked up again on every start: the browser may have been updated or reinstalled.
		b, err := findBrowser(p.browser)
		if err != nil {
			return err
		}
		started := time.Now()
		cmd := kioskCommand(b, p.address)
		p.mu.Lock()
		if p.quitting {
			p.mu.Unlock()
			return nil
		}
		p.current = cmd
		p.mu.Unlock()
		log.Printf("starting %s in kiosk mode: %s", b.Name, p.address)
		if err := cmd.Run(); err != nil {
			log.Printf("%s exited: %v", b.Name, err)
		} else {
			log.Printf("%s closed after %s", b.Name, time.Since(started).Round(time.Second))
		}
		p.mu.Lock()
		stop, restart := p.quitting, p.restarting
		p.restarting = false
		p.mu.Unlock()
		if stop {
			return nil
		}
		if restart {
			continue
		}

		now := time.Now()
		exits = append(exits, now)
		for len(exits) > 0 && now.Sub(exits[0]) > quickExitWindow {
			exits = exits[1:]
		}
		if len(exits) >= maxQuickExits {
			log.Printf("browser closed %d times in %s: stopping", len(exits), quickExitWindow)
			return nil
		}
		// A browser that ran for a while crashed or was killed: come back quickly. One that died at
		// once (no network, a broken update) gets an increasing pause.
		if time.Since(started) > 10*time.Minute {
			delay = 2 * time.Second
		} else if delay < time.Minute {
			delay *= 2
		}
		time.Sleep(delay)
	}
}

// stopPresenting asks a running "launch" to close its kiosk and exit. It reports whether one was
// running, so the updater knows to start the presentation again afterwards.
func stopPresenting() bool {
	conn, err := net.DialTimeout("tcp", presentLock, time.Second)
	if err != nil {
		return false
	}
	defer conn.Close()
	_, _ = conn.Write([]byte("quit\n"))
	// The port frees when that process exits.
	for i := 0; i < 50; i++ {
		time.Sleep(200 * time.Millisecond)
		if l, err := net.Listen("tcp", presentLock); err == nil {
			l.Close()
			break
		}
	}
	return true
}

// offerUpdateAtBoot shows the update prompt when a newer version exists and installs it if the
// person accepts. It returns true when the installer took over (it starts the presentation again).
// release frees the presentation lock first, so the launcher the update starts can take it.
func offerUpdateAtBoot(release func()) bool {
	ctx, cancel := context.WithTimeout(context.Background(), 8*time.Second)
	u, err := availableUpdate(ctx)
	cancel()
	if err != nil {
		log.Printf("update check: %v", err)
		return false
	}
	if u == nil || !promptUpdate(u) {
		return false
	}
	path, err := download(context.Background(), u)
	if err != nil {
		log.Printf("update download: %v", err)
		return false
	}
	release()
	if err := installUpdate(path, "launch"); err != nil {
		log.Printf("update install: %v", err)
		return false
	}
	log.Printf("updating to %s", u.Version)
	return true
}
