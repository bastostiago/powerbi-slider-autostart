// Power Slider Launcher: opens a Power Slider presentation fullscreen when the computer starts,
// and opens it again if the browser crashes.
//
//	launcher          the settings window
//	launcher launch   the presentation, in kiosk mode, watched (what the boot shortcut runs)
//
// One process does everything (see app.go): a second start only tells the running one what to do.
package main

import (
	"bufio"
	"context"
	"errors"
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

// instanceLock is held by the one running launcher. A loopback port is the lock: the OS frees it
// when the process dies, however it dies. It is also how a second start talks to the first one:
// "show" (the settings window), "start" (the presentation) or "quit".
var instanceLock = func() string {
	if os.Getenv("PSL_DATA_DIR") != "" {
		return "127.0.0.1:47614" // a trial run must not meet the installed launcher
	}
	return "127.0.0.1:47613"
}()

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
	launch := len(os.Args) > 1 && os.Args[1] == "launch"

	lock, err := net.Listen("tcp", instanceLock)
	if err != nil {
		command := "show"
		if launch {
			command = "start"
		}
		if sendCommand(command) {
			log.Printf("already running: sent %q", command)
			return
		}
		log.Printf("instance lock busy and nobody answers: %v", err)
		os.Exit(1)
	}

	a := newApp(lock)
	if launch {
		// At boot someone may be in front of the screen: offer a pending update before presenting.
		// An unanswered prompt times out and the presentation goes on.
		if offerUpdateAtBoot(func() { lock.Close() }) {
			return
		}
		if err := a.startPresentation(); err != nil {
			log.Printf("launch: %v", err)
		}
	} else if err := a.showSettings(); err != nil {
		log.Printf("settings: %v", err)
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
	a.run()
}

// sendCommand hands a command to the running launcher.
func sendCommand(command string) bool {
	conn, err := net.DialTimeout("tcp", instanceLock, time.Second)
	if err != nil {
		return false
	}
	defer conn.Close()
	_, err = conn.Write([]byte(command + "\n"))
	return err == nil
}

// readCommand reads one control command from a connection.
func readCommand(conn net.Conn) string {
	_ = conn.SetReadDeadline(time.Now().Add(5 * time.Second))
	line, _ := bufio.NewReader(conn).ReadString('\n')
	return line
}

// Crash limits: a browser that crashes this many times inside the window cannot start (no
// network, a broken update), and the launcher stops trying.
const (
	maxCrashes  = 3
	crashWindow = 2 * time.Minute
)

// presenter is the watchdog: it keeps the kiosk browser open until the presentation is stopped.
type presenter struct {
	browser string
	address string

	mu         sync.Mutex
	current    *exec.Cmd
	quitting   bool
	restarting bool
}

// stop closes the browser and ends the watchdog.
func (p *presenter) stop() {
	p.mu.Lock()
	defer p.mu.Unlock()
	p.quitting = true
	p.killCurrent()
}

// restart closes the browser and opens it again at once.
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

// watch runs the browser until the presentation ends. A browser that a person closed exits with
// code 0 (Alt+F4, the window's X, Windows shutting down) and ends it; one that crashed or was
// killed exits with an error code and is opened again.
func (p *presenter) watch() error {
	var crashes []time.Time
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
		runErr := cmd.Run()

		p.mu.Lock()
		stop, restart := p.quitting, p.restarting
		p.restarting = false
		p.mu.Unlock()
		switch {
		case stop:
			log.Printf("presentation stopped")
			return nil
		case restart:
			continue
		case runErr == nil:
			log.Printf("%s closed by the user after %s", b.Name, time.Since(started).Round(time.Second))
			return nil
		}

		var exitErr *exec.ExitError
		if !errors.As(runErr, &exitErr) {
			return runErr // could not start at all
		}
		log.Printf("%s crashed after %s: %v", b.Name, time.Since(started).Round(time.Second), runErr)
		now := time.Now()
		crashes = append(crashes, now)
		for len(crashes) > 0 && now.Sub(crashes[0]) > crashWindow {
			crashes = crashes[1:]
		}
		if len(crashes) >= maxCrashes {
			return fmt.Errorf("browser crashed %d times in %s", len(crashes), crashWindow)
		}
		// A browser that ran for a while comes back quickly; one that dies at once waits longer.
		if time.Since(started) > 10*time.Minute {
			delay = 2 * time.Second
		} else if delay < time.Minute {
			delay *= 2
		}
		time.Sleep(delay)
	}
}

// offerUpdateAtBoot shows the update prompt when a newer version exists and installs it if the
// person accepts. It returns true when the installer took over (it starts the presentation again).
// release frees the instance lock first, so the launcher the update starts can take it.
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
