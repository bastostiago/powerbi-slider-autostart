package main

import (
	"context"
	"crypto/sha256"
	"fmt"
	"net/http"
	"net/http/httptest"
	"os"
	"strings"
	"testing"
)

func TestNewer(t *testing.T) {
	cases := []struct {
		a, b string
		want bool
	}{
		{"3.0.1", "3.0.0", true},
		{"3.1.0", "3.0.9", true},
		{"10.0.0", "9.9.9", true},
		{"3.0.0", "3.0.0", false},
		{"2.9.9", "3.0.0", false},
		{"3.0.1", "dev", false},
	}
	for _, c := range cases {
		if got := newer(c.a, c.b); got != c.want {
			t.Errorf("newer(%q, %q) = %v; want %v", c.a, c.b, got, c.want)
		}
	}
}

func TestAvailableUpdateAndDownload(t *testing.T) {
	payload := []byte("installer bytes")
	sum := sha256.Sum256(payload)
	var srv *httptest.Server
	srv = httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch r.URL.Path {
		case "/latest":
			fmt.Fprintf(w, `{"tag_name":"v3.1.0","assets":[{"name":%q,"browser_download_url":%q,"digest":"sha256:%x"},{"name":"other.zip","browser_download_url":"x"}]}`,
				updateAssetName(), srv.URL+"/asset", sum)
		case "/asset":
			w.Write(payload)
		}
	}))
	defer srv.Close()
	oldAPI, oldVersion := releasesAPI, version
	releasesAPI, version = srv.URL+"/latest", "3.0.0"
	defer func() { releasesAPI, version = oldAPI, oldVersion }()

	u, err := availableUpdate(context.Background())
	if err != nil || u == nil || u.Version != "3.1.0" {
		t.Fatalf("availableUpdate = %+v, %v", u, err)
	}
	path, err := download(context.Background(), u)
	if err != nil {
		t.Fatal(err)
	}
	defer os.Remove(path)
	if got, _ := os.ReadFile(path); string(got) != string(payload) {
		t.Fatalf("downloaded %q", got)
	}

	u.Digest = "sha256:" + strings.Repeat("0", 64)
	if _, err := download(context.Background(), u); err == nil {
		t.Fatal("a wrong checksum was accepted")
	}

	version = "3.1.0"
	if u, err := availableUpdate(context.Background()); err != nil || u != nil {
		t.Fatalf("current version offered an update: %+v, %v", u, err)
	}
}
