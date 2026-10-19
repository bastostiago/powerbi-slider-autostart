package main

import (
	"context"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"
)

// Releases are published on GitHub with fixed asset names (see build.ps1), so the updater and the
// site's download page both find the right file under "latest".
var releasesAPI = "https://api.github.com/repos/bastostiago/powerbi-slider-autostart/releases/latest"

type update struct {
	Version string
	Asset   string
	URL     string
	Digest  string // "sha256:<hex>", as GitHub reports it; empty on assets uploaded before it did
}

// availableUpdate returns the newer release for this platform, or nil when this one is current.
func availableUpdate(ctx context.Context) (*update, error) {
	req, err := http.NewRequestWithContext(ctx, http.MethodGet, releasesAPI, nil)
	if err != nil {
		return nil, err
	}
	req.Header.Set("Accept", "application/vnd.github+json")
	req.Header.Set("User-Agent", "power-slider-launcher/"+version)
	res, err := http.DefaultClient.Do(req)
	if err != nil {
		return nil, err
	}
	defer res.Body.Close()
	if res.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("releases: HTTP %d", res.StatusCode)
	}
	var rel struct {
		TagName string `json:"tag_name"`
		Assets  []struct {
			Name   string `json:"name"`
			URL    string `json:"browser_download_url"`
			Digest string `json:"digest"`
		} `json:"assets"`
	}
	if err := json.NewDecoder(res.Body).Decode(&rel); err != nil {
		return nil, err
	}
	latest := strings.TrimPrefix(rel.TagName, "v")
	if !newer(latest, version) {
		return nil, nil
	}
	name := updateAssetName()
	for _, a := range rel.Assets {
		if a.Name == name {
			return &update{Version: latest, Asset: a.Name, URL: a.URL, Digest: a.Digest}, nil
		}
	}
	return nil, nil // released without a package for this platform
}

// newer reports whether a is a later version than b ("3.1.0" > "3.0.9"). A development build is
// never offered an update.
func newer(a, b string) bool {
	if b == "dev" {
		return false
	}
	pa, pb := versionParts(a), versionParts(b)
	for i := 0; i < 3; i++ {
		if pa[i] != pb[i] {
			return pa[i] > pb[i]
		}
	}
	return false
}

func versionParts(v string) [3]int {
	var out [3]int
	for i, part := range strings.SplitN(v, ".", 3) {
		n, _ := strconv.Atoi(strings.TrimFunc(part, func(r rune) bool { return r < '0' || r > '9' }))
		out[i] = n
	}
	return out
}

// download saves the package to the temp folder and checks it against GitHub's digest.
func download(ctx context.Context, u *update) (string, error) {
	ctx, cancel := context.WithTimeout(ctx, 10*time.Minute)
	defer cancel()
	req, err := http.NewRequestWithContext(ctx, http.MethodGet, u.URL, nil)
	if err != nil {
		return "", err
	}
	req.Header.Set("User-Agent", "power-slider-launcher/"+version)
	res, err := http.DefaultClient.Do(req)
	if err != nil {
		return "", err
	}
	defer res.Body.Close()
	if res.StatusCode != http.StatusOK {
		return "", fmt.Errorf("download: HTTP %d", res.StatusCode)
	}
	path := filepath.Join(os.TempDir(), u.Asset)
	f, err := os.Create(path)
	if err != nil {
		return "", err
	}
	hash := sha256.New()
	_, err = io.Copy(io.MultiWriter(f, hash), res.Body)
	if cerr := f.Close(); err == nil {
		err = cerr
	}
	if err != nil {
		return "", err
	}
	if want, ok := strings.CutPrefix(u.Digest, "sha256:"); ok {
		if got := hex.EncodeToString(hash.Sum(nil)); !strings.EqualFold(got, want) {
			os.Remove(path)
			return "", errors.New("download: checksum mismatch")
		}
	}
	return path, nil
}
