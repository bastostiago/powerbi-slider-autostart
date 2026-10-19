<h1>Power Slider Launcher</h1>
<p>Opens a <a href="https://www.powerslider.app">Power Slider</a> presentation in full screen when the computer starts, and opens it again if the browser closes or crashes. Windows and Linux (Debian, Ubuntu, Raspberry Pi OS), with Google Chrome, Microsoft Edge or Chromium.</p>

<h2>How it works</h2>
<ul>
  <li><code>launcher</code> (no arguments) opens the settings window: a local page shown as an Edge/Chrome app window, in the site's eight languages. Save the presentation address, sign in once with Microsoft, start.</li>
  <li><code>launcher launch</code> is what the boot shortcut runs: the browser in kiosk mode on the address plus <code>?present=1</code> (the site starts the slideshow without a click), restarted when it closes. Three closes within two minutes stop it, so a person can get out with Alt+F4.</li>
  <li>Updates come from this repository's latest GitHub release. The launcher asks before installing: in the settings window, and at boot (a prompt that gives up after a minute so an unattended screen still presents).</li>
</ul>
<p>Version 3 replaced the Python/Selenium launcher. It keeps its executable name, Startup shortcut (<code>PB Slider Launcher.lnk</code> with <code>launch</code>), install folder and data folder (<code>%LOCALAPPDATA%\PB Slider Launcher</code>, with <code>config.ini</code> and the browser profiles), so a screen updated in place keeps its address and its Microsoft session.</p>

<h2>Building</h2>
<p>Needs Go, Inno Setup 6 and Python 3. From the repo root:</p>
<pre>.\build.ps1</pre>
<p>Everything lands in <code>dist\release</code>, with fixed names that the updater and the site's download page look up under <code>releases/latest/download/</code>:</p>
<ul>
  <li><code>Power_Slider_Launcher_Setup.exe</code>: Windows installer</li>
  <li><code>power-slider-launcher_amd64.deb</code>, <code>power-slider-launcher_arm64.deb</code>: Debian, Ubuntu, Raspberry Pi OS</li>
  <li><code>power-slider-launcher-linux-amd64.tar.gz</code>, <code>power-slider-launcher-linux-arm64.tar.gz</code>: other distributions</li>
  <li><code>PB_Slider_Launcher_Installer.zip</code>: the installer under its 1.x name, for old links</li>
</ul>
<p>The version comes from <code>VERSION</code>. To release: bump it, run the build, and publish a GitHub release tagged <code>v&lt;version&gt;</code> with every file in <code>dist\release</code>.</p>

<h2>Code</h2>
<p>All in <code>launcher/</code>: <code>main.go</code> (modes and the watchdog), <code>settings.go</code> (the local windows), <code>update.go</code>, <code>config.go</code>, <code>browser.go</code>, and the platform files. <code>go test ./...</code> runs on both platforms.</p>
