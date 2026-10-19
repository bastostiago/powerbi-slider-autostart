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
	"net/url"
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
	ui := &localUI{mux: http.NewServeMux(), base: "/" + hex.EncodeToString(token), ln: ln}
	ui.mux.Handle(ui.base+"/", http.StripPrefix(ui.base, http.FileServerFS(mustSub(uiFiles, "ui"))))
	ui.server = &http.Server{Handler: ui.mux}
	go ui.server.Serve(ln)
	return ui, nil
}

func (ui *localUI) handle(name string, h http.HandlerFunc) {
	ui.mux.HandleFunc(ui.base+"/api/"+name, h)
}

func (ui *localUI) open(page string, width, height int) (*exec.Cmd, error) {
	b, err := findBrowser("auto")
	if err != nil {
		return nil, err
	}
	// The pages speak the operating system's language, not the browser's (they can differ, and the
	// tray menu already follows the system).
	address := "http://" + ui.ln.Addr().String() + ui.base + "/" + page + "?lang=" + url.QueryEscape(normalizeLocale(systemLanguage()))
	cmd := appWindowCommand(b, address, width, height)
	return cmd, cmd.Start()
}

type state struct {
	URL        string   `json:"url"`
	Browser    string   `json:"browser"`
	Autostart  bool     `json:"autostart"`
	Browsers   []string `json:"browsers"`
	Version    string   `json:"version"`
	Presenting bool     `json:"presenting"`
}

// registerSettings wires the settings window to the running launcher.
func (a *app) registerSettings(ui *localUI) {
	ui.handle("ping", func(w http.ResponseWriter, r *http.Request) {
		a.ping()
		w.WriteHeader(http.StatusNoContent)
	})
	ui.handle("state", func(w http.ResponseWriter, r *http.Request) {
		a.ping()
		cfg := loadConfig()
		writeJSON(w, state{URL: cfg.URL, Browser: cfg.Browser, Autostart: autostartEnabled(), Browsers: installedBrowsers(), Version: version, Presenting: a.presenting()})
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
	// Starts (or switches to) the saved presentation and closes the window: the launcher goes on in
	// the tray, where Settings brings the window back. The first Microsoft sign-in happens inside
	// the presentation itself.
	ui.handle("start", func(w http.ResponseWriter, r *http.Request) {
		if err := a.startPresentation(); err != nil {
			httpError(w, "start_failed", err)
			return
		}
		writeJSON(w, map[string]bool{"ok": true})
		go func() {
			time.Sleep(300 * time.Millisecond)
			a.closeWindow()
		}()
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
	// Installs the update the window just showed. This launcher closes completely first (window,
	// presentation, process): the installer replaces its executable and starts the new version on
	// whatever was open.
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
		if a.presenting() {
			relaunch = "launch"
		}
		writeJSON(w, map[string]bool{"ok": true})
		go func() {
			a.stopPresentation()
			a.lock.Close() // the new version takes the instance lock
			if err := installUpdate(path, relaunch); err != nil {
				log.Printf("update install: %v", err)
			} else {
				log.Printf("updating to %s", u.Version)
			}
			a.closeWindow()
			a.exit()
		}()
	})
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
	ui.handle("ping", func(w http.ResponseWriter, r *http.Request) { w.WriteHeader(http.StatusNoContent) })
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
