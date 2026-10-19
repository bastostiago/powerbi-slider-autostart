package main

import (
	"context"
	"crypto/rand"
	"embed"
	"encoding/hex"
	"encoding/json"
	"errors"
	"log"
	"net"
	"net/http"
	"os"
	"os/exec"
	"sync"
	"time"
)

//go:embed ui
var uiFiles embed.FS

// localUI serves the embedded pages on 127.0.0.1 and shows one of them in an app window of
// Edge/Chrome, so the same windows work on Windows and Linux without a GUI toolkit. Every route sits
// under a random token, so no web page can drive the launcher through the local port.
type localUI struct {
	mux    *http.ServeMux
	base   string
	server *http.Server
	ln     net.Listener

	mu       sync.Mutex
	lastPing time.Time
}

func newLocalUI() (*localUI, error) {
	token := make([]byte, 16)
	if _, err := rand.Read(token); err != nil {
		return nil, err
	}
	ln, err := net.Listen("tcp", "127.0.0.1:0")
	if err != nil {
		return nil, err
	}
	ui := &localUI{mux: http.NewServeMux(), base: "/" + hex.EncodeToString(token), ln: ln, lastPing: time.Now()}
	ui.mux.Handle(ui.base+"/", http.StripPrefix(ui.base, http.FileServerFS(mustSub(uiFiles, "ui"))))
	ui.handle("ping", func(w http.ResponseWriter, r *http.Request) { w.WriteHeader(http.StatusNoContent) })
	ui.server = &http.Server{Handler: ui.mux}
	go ui.server.Serve(ln)
	return ui, nil
}

// handle registers an API route; every call counts as a heartbeat from the window.
func (ui *localUI) handle(name string, h http.HandlerFunc) {
	ui.mux.HandleFunc(ui.base+"/api/"+name, func(w http.ResponseWriter, r *http.Request) {
		ui.mu.Lock()
		ui.lastPing = time.Now()
		ui.mu.Unlock()
		h(w, r)
	})
}

func (ui *localUI) idle() time.Duration {
	ui.mu.Lock()
	defer ui.mu.Unlock()
	return time.Since(ui.lastPing)
}

func (ui *localUI) open(page string, width, height int) (*exec.Cmd, error) {
	b, err := findBrowser("auto")
	if err != nil {
		return nil, err
	}
	cmd := appWindowCommand(b, "http://"+ui.ln.Addr().String()+ui.base+"/"+page, width, height)
	return cmd, cmd.Start()
}

type state struct {
	URL       string   `json:"url"`
	Browser   string   `json:"browser"`
	Autostart bool     `json:"autostart"`
	Browsers  []string `json:"browsers"`
	Version   string   `json:"version"`
}

// runSettings shows the settings window. The process ends when the window stops sending heartbeats.
func runSettings() error {
	ui, err := newLocalUI()
	if err != nil {
		return err
	}
	ui.handle("state", func(w http.ResponseWriter, r *http.Request) {
		cfg := loadConfig()
		writeJSON(w, state{URL: cfg.URL, Browser: cfg.Browser, Autostart: autostartEnabled(), Browsers: installedBrowsers(), Version: version})
	})
	ui.handle("save", func(w http.ResponseWriter, r *http.Request) {
		var in state
		if err := readJSON(r, &in); err != nil {
			httpError(w, "bad_request", err)
			return
		}
		address, ok := normalizeURL(in.URL)
		if !ok {
			httpError(w, "invalid_url", errors.New("not a Power Slider URL"))
			return
		}
		if err := saveConfig(Config{URL: address, Browser: in.Browser}); err != nil {
			httpError(w, "save_failed", err)
			return
		}
		if err := setAutostart(in.Autostart); err != nil {
			httpError(w, "autostart_failed", err)
			return
		}
		writeJSON(w, map[string]string{"url": address})
	})
	// Opens the saved address in a normal window of the kiosk's own profile: the Microsoft sign-in
	// has to happen there once, and a kiosk window cannot be closed by a person.
	ui.handle("signin", func(w http.ResponseWriter, r *http.Request) {
		cfg := loadConfig()
		address, ok := normalizeURL(cfg.URL)
		if !ok {
			httpError(w, "invalid_url", errors.New("save a Power Slider URL first"))
			return
		}
		kb, err := findBrowser(cfg.Browser)
		if err != nil {
			httpError(w, "no_browser", err)
			return
		}
		if err := windowCommand(kb, address).Start(); err != nil {
			httpError(w, "start_failed", err)
			return
		}
		writeJSON(w, map[string]bool{"ok": true})
	})
	// Starts the presentation exactly as the boot shortcut does, in a process of its own.
	ui.handle("start", func(w http.ResponseWriter, r *http.Request) {
		if err := startDetached("launch"); err != nil {
			httpError(w, "start_failed", err)
			return
		}
		writeJSON(w, map[string]bool{"ok": true})
	})

	var (
		pendingMu sync.Mutex
		pending   *update
	)
	ui.handle("update", func(w http.ResponseWriter, r *http.Request) {
		ctx, cancel := context.WithTimeout(r.Context(), 10*time.Second)
		defer cancel()
		u, err := availableUpdate(ctx)
		if err != nil {
			log.Printf("update check: %v", err)
		}
		pendingMu.Lock()
		pending = u
		pendingMu.Unlock()
		if u == nil {
			writeJSON(w, map[string]any{"available": false, "current": version})
			return
		}
		writeJSON(w, map[string]any{"available": true, "current": version, "latest": u.Version})
	})
	// Installs the update the window just showed. A running presentation is stopped first (the
	// installer replaces the executable it runs from) and started again by the new version.
	ui.handle("update/apply", func(w http.ResponseWriter, r *http.Request) {
		pendingMu.Lock()
		u := pending
		pendingMu.Unlock()
		if u == nil {
			httpError(w, "no_update", errors.New("no update to apply"))
			return
		}
		path, err := download(r.Context(), u)
		if err != nil {
			httpError(w, "download_failed", err)
			return
		}
		relaunch := "settings"
		if stopPresenting() {
			relaunch = "launch"
		}
		if err := installUpdate(path, relaunch); err != nil {
			httpError(w, "install_failed", err)
			return
		}
		log.Printf("updating to %s", u.Version)
		writeJSON(w, map[string]bool{"ok": true})
		go func() { time.Sleep(time.Second); os.Exit(0) }()
	})

	if _, err := ui.open("index.html", 560, 680); err != nil {
		return err
	}
	for {
		time.Sleep(5 * time.Second)
		if ui.idle() > 20*time.Second {
			return ui.server.Close()
		}
	}
}

// promptUpdate asks, in a small window, whether to install the update now. No answer within a
// minute counts as "not now": an unattended screen must get to its presentation.
func promptUpdate(u *update) bool {
	ui, err := newLocalUI()
	if err != nil {
		log.Printf("update prompt: %v", err)
		return false
	}
	defer ui.server.Close()
	answer := make(chan bool, 1)
	ui.handle("info", func(w http.ResponseWriter, r *http.Request) {
		writeJSON(w, map[string]string{"current": version, "latest": u.Version})
	})
	ui.handle("answer", func(w http.ResponseWriter, r *http.Request) {
		var in struct {
			Update bool `json:"update"`
		}
		_ = readJSON(r, &in)
		writeJSON(w, map[string]bool{"ok": true})
		select {
		case answer <- in.Update:
		default:
		}
	})
	window, err := ui.open("update.html", 460, 300)
	if err != nil {
		log.Printf("update prompt: %v", err)
		return false
	}
	defer func() {
		if window.Process != nil {
			_ = window.Process.Kill()
		}
	}()
	select {
	case yes := <-answer:
		log.Printf("update %s: answered %v", u.Version, yes)
		return yes
	case <-time.After(60 * time.Second):
		log.Printf("update %s: no answer, presenting", u.Version)
		return false
	}
}

func startDetached(args ...string) error {
	exe, err := os.Executable()
	if err != nil {
		return err
	}
	cmd := exec.Command(exe, args...)
	hideWindow(cmd)
	return cmd.Start()
}

func installedBrowsers() []string {
	var names []string
	for _, name := range []string{"chrome", "edge", "chromium"} {
		if _, err := findBrowser(name); err == nil {
			names = append(names, name)
		}
	}
	return names
}

func readJSON(r *http.Request, v any) error {
	if r.Method != http.MethodPost {
		return errors.New("POST only")
	}
	return json.NewDecoder(http.MaxBytesReader(nil, r.Body, 1<<16)).Decode(v)
}

func writeJSON(w http.ResponseWriter, v any) {
	w.Header().Set("Content-Type", "application/json")
	_ = json.NewEncoder(w).Encode(v)
}

func httpError(w http.ResponseWriter, code string, err error) {
	log.Printf("%s: %v", code, err)
	w.Header().Set("Content-Type", "application/json")
	w.WriteHeader(http.StatusBadRequest)
	_ = json.NewEncoder(w).Encode(map[string]string{"error": code})
}
