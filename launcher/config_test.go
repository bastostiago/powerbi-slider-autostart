package main

import (
	"os"
	"path/filepath"
	"testing"
)

func TestNormalizeURL(t *testing.T) {
	cases := map[string]string{
		"https://www.powerslider.app/en/home/me/report/r1":   "https://www.powerslider.app/en/home/me/report/r1",
		"www.powerbislider.com/pt-BR/home/g/report/r1":       "https://www.powerbislider.com/pt-BR/home/g/report/r1",
		"  https://powerslider.app/en/home/composition/42  ": "https://powerslider.app/en/home/composition/42",
		"http://localhost:3000/en/home/me/report/r1":         "http://localhost:3000/en/home/me/report/r1",
	}
	for in, want := range cases {
		if got, ok := normalizeURL(in); !ok || got != want {
			t.Errorf("normalizeURL(%q) = %q, %v; want %q", in, got, ok, want)
		}
	}
	for _, in := range []string{"", "https://evil.example/powerslider.app", "https://powerslider.app.evil.com/", "https://app.powerbi.com/groups/me/reports/r1"} {
		if got, ok := normalizeURL(in); ok {
			t.Errorf("normalizeURL(%q) accepted as %q", in, got)
		}
	}
}

func TestPresentationURL(t *testing.T) {
	cases := map[string]string{
		"https://www.powerslider.app/en/home/me/report/r1?openSettings=1": "https://www.powerslider.app/en/home/me/report/r1?present=1",
		"https://www.powerslider.app/en/home/composition/42":              "https://www.powerslider.app/en/home/composition/42?present=1",
		"https://www.powerslider.app/en/home":                             "https://www.powerslider.app/en/home",
	}
	for in, want := range cases {
		if got := presentationURL(in); got != want {
			t.Errorf("presentationURL(%q) = %q; want %q", in, got, want)
		}
	}
}

// A config.ini written by the Python launcher must still be read.
func TestLoadsOldConfig(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("LOCALAPPDATA", dir)
	t.Setenv("XDG_CONFIG_HOME", dir)
	if err := os.MkdirAll(dataDir(), 0o755); err != nil {
		t.Fatal(err)
	}
	old := "[DEFAULT]\nurl=https://www.powerbislider.com/en/home/g/report/r1"
	if err := os.WriteFile(filepath.Join(dataDir(), "config.ini"), []byte(old), 0o644); err != nil {
		t.Fatal(err)
	}
	cfg := loadConfig()
	if cfg.URL != "https://www.powerbislider.com/en/home/g/report/r1" || cfg.Browser != "auto" {
		t.Fatalf("got %+v", cfg)
	}
	cfg.Browser = "edge"
	if err := saveConfig(cfg); err != nil {
		t.Fatal(err)
	}
	if again := loadConfig(); again != cfg {
		t.Fatalf("round trip: got %+v, want %+v", again, cfg)
	}
}

func TestTrayText(t *testing.T) {
	cases := map[string]string{
		"pt-BR":       "Parar apresentação",
		"pt_PT.UTF-8": "Parar apresentação",
		"de-AT":       "Präsentation beenden",
		"es_419":      "Detener presentación",
		"ja-JP":       "Stop presentation",
		"":            "Stop presentation",
	}
	for locale, want := range cases {
		if got := trayText(locale).stop; got != want {
			t.Errorf("trayText(%q).stop = %q; want %q", locale, got, want)
		}
	}
}

func TestNormalizeLocale(t *testing.T) {
	for in, want := range map[string]string{"pt_BR.UTF-8": "pt-BR", "de_AT@euro": "de-AT", "fr-FR": "fr-FR", "en": "en"} {
		if got := normalizeLocale(in); got != want {
			t.Errorf("normalizeLocale(%q) = %q; want %q", in, got, want)
		}
	}
}
