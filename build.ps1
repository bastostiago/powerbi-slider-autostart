# Builds every release asset into dist\release. Needs Go, Inno Setup 6 and Python 3 (for the
# tarballs, so the binary keeps its executable bit). Run from the repo root: .\build.ps1
#
# The asset names are fixed on purpose: the launcher's updater and the site's download page find
# them under https://github.com/bastostiago/powerbi-slider-autostart/releases/latest/download/<name>.
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$version = (Get-Content VERSION -Raw).Trim()
$out = 'dist\release'
$build = 'dist\build'
Remove-Item -Recurse -Force $out, $build -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force $out, $build | Out-Null
$ldflags = "-s -w -X main.version=$version"

function Invoke-Checked([scriptblock]$block) {
  & $block
  if ($LASTEXITCODE -ne 0) { throw "failed: $block" }
}

Push-Location launcher
try {
  Invoke-Checked { go test ./... }
  # Icon and version information for the Windows executable.
  Invoke-Checked { go run github.com/tc-hib/go-winres@v0.3.3 simply --icon ..\logo.png --manifest gui --product-version $version --file-version $version --product-name 'Power Slider Launcher' --file-description 'Power Slider Launcher' --original-filename 'PB Slider Launcher.exe' }
  $env:GOOS = 'windows'; $env:GOARCH = 'amd64'
  Invoke-Checked { go build -trimpath -ldflags "$ldflags -H windowsgui" -o "..\$build\windows-amd64\PB Slider Launcher.exe" . }
  foreach ($arch in 'amd64', 'arm64') {
    $env:GOOS = 'linux'; $env:GOARCH = $arch
    Invoke-Checked { go build -trimpath -ldflags $ldflags -o "..\$build\linux-$arch\power-slider-launcher" . }
  }
} finally {
  Remove-Item Env:GOOS, Env:GOARCH -ErrorAction SilentlyContinue
  Remove-Item rsrc_windows_*.syso -ErrorAction SilentlyContinue
  Pop-Location
}

# Linux: .deb (Debian, Ubuntu, Raspberry Pi OS) and a tarball for everything else.
foreach ($arch in 'amd64', 'arm64') {
  $env:VERSION = $version; $env:ARCH = $arch
  New-Item -ItemType Directory -Force "$build\deb" | Out-Null
  Copy-Item "$build\linux-$arch\power-slider-launcher" "$build\deb\power-slider-launcher" -Force
  Invoke-Checked { go run github.com/goreleaser/nfpm/v2/cmd/nfpm@v2.43.0 package --config packaging\nfpm.yaml --packager deb --target "$out\power-slider-launcher_$arch.deb" }
  Invoke-Checked { python -c "import sys, tarfile; t = tarfile.open(sys.argv[1], 'w:gz'); i = t.gettarinfo(sys.argv[2], 'power-slider-launcher'); i.mode = 0o755; i.uid = i.gid = 0; i.uname = i.gname = 'root'; t.addfile(i, open(sys.argv[2], 'rb')); t.close()" "$out\power-slider-launcher-linux-$arch.tar.gz" "$build\linux-$arch\power-slider-launcher" }
}
Remove-Item Env:VERSION, Env:ARCH -ErrorAction SilentlyContinue

# Windows installer.
$iscc = @("$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe", "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) { throw 'Inno Setup 6 not found' }
Invoke-Checked { & $iscc /Q "/DAppVersion=$version" inno_setup.iss }

# The 1.x name, zipped: the site linked to it, so an old link keeps downloading the current version.
Compress-Archive -Path "$out\Power_Slider_Launcher_Setup.exe" -DestinationPath "$out\PB_Slider_Launcher_Installer.zip" -Force

Get-ChildItem $out | ForEach-Object { '{0}  {1}' -f (Get-FileHash $_.FullName -Algorithm SHA256).Hash.ToLower(), $_.Name } | Set-Content "$out\SHA256SUMS.txt"
Get-ChildItem $out | Select-Object Name, Length
