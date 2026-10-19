package main

import (
	"bufio"
	"fmt"
	"net/url"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
)

// Config is config.ini. The 1.x/2.x Python launcher wrote it as "[DEFAULT]\nurl=...", and a screen
// updated in place must keep its URL, so the format stays an INI with a [DEFAULT] section.
type Config struct {
	URL     string
	Browser string // "auto", "chrome", "edge" or "chromium"
}

func configPath() string { return filepath.Join(dataDir(), "config.ini") }

func loadConfig() Config {
	cfg := Config{Browser: "auto"}
	f, err := os.Open(configPath())
	if err != nil {
		return cfg
	}
	defer f.Close()
	scanner := bufio.NewScanner(f)
	for scanner.Scan() {
		key, value, ok := strings.Cut(scanner.Text(), "=")
		if !ok {
			continue
		}
		switch strings.TrimSpace(strings.ToLower(key)) {
		case "url":
			cfg.URL = strings.TrimSpace(value)
		case "browser":
			cfg.Browser = strings.TrimSpace(value)
		}
	}
	return cfg
}

func saveConfig(cfg Config) error {
	if err := os.MkdirAll(dataDir(), 0o755); err != nil {
		return err
	}
	values := map[string]string{"url": cfg.URL, "browser": cfg.Browser}
	keys := make([]string, 0, len(values))
	for k := range values {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	var b strings.Builder
	b.WriteString("[DEFAULT]\n")
	for _, k := range keys {
		fmt.Fprintf(&b, "%s=%s\n", k, values[k])
	}
	return os.WriteFile(configPath(), []byte(b.String()), 0o644)
}

// The site moved from powerbislider.com to powerslider.app and serves both.
var siteURL = regexp.MustCompile(`^https?://(localhost(:\d+)?|(www\.)?powerslider\.app|(www\.)?powerbislider\.com)(/.*)?$`)

// normalizeURL validates a Power Slider address and returns it with https:// when the scheme was
// left out. ok is false for anything that is not the site.
func normalizeURL(raw string) (string, bool) {
	raw = strings.TrimSpace(raw)
	if raw == "" {
		return "", false
	}
	if !strings.Contains(raw, "://") {
		raw = "https://" + raw
	}
	if !siteURL.MatchString(raw) {
		return "", false
	}
	if _, err := url.Parse(raw); err != nil {
		return "", false
	}
	return raw, true
}

// presentationURL adds present=1, which makes the site start the slideshow without a click (a
// kiosk has nobody to click) on report and composition pages.
func presentationURL(raw string) string {
	u, err := url.Parse(raw)
	if err != nil {
		return raw
	}
	if !strings.Contains(u.Path, "/report/") && !strings.Contains(u.Path, "/composition/") {
		return raw
	}
	q := u.Query()
	q.Set("present", "1")
	q.Del("openSettings")
	u.RawQuery = q.Encode()
	return u.String()
}
