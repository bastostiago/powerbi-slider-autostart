<h1>Power Slider Launcher</h1>
<p>Opens a <a href="https://www.powerslider.app">Power Slider</a> presentation in full screen when the computer starts, and opens it again if the browser closes or crashes. Windows and Linux (Debian, Ubuntu, Raspberry Pi OS), with Google Chrome, Microsoft Edge or Chromium.</p>

<h2>How it works</h2>
<p>One process does everything: the presentation, the settings window and an icon in the notification area. Starting the launcher again only talks to the running one (it brings the window to the front, or starts the presentation).</p>
<ul>
  <li><code>launcher</code> (no arguments) opens the settings window: a local page shown as an Edge/Chrome app window, in the operating system's language (the site's eight; English otherwise). Paste the address and <strong>Start presentation</strong>: the window closes, the presentation opens and the launcher stays in the notification area. The first Microsoft sign-in happens inside the presentation.</li>
  <li><code>launcher launch</code> is what the boot shortcut runs: the browser in kiosk mode on the address plus <code>?present=1</code> (the site starts the slideshow without a click).</li>
  <li>The tray menu stops, restarts or starts the presentation and opens the settings (the Windows key brings the taskbar over the kiosk). On Linux it needs a desktop that shows tray icons (StatusNotifierItem); without one the launcher runs without the icon.</li>
  <li>A browser a person closes (exit code 0) ends the presentation; one that crashes is opened again, up to three times in two minutes. With no presentation and no window, the launcher exits.</li>
  <li>Updates come from this repository's latest GitHub release. The launcher asks before installing: in the settings window, and at boot (a prompt that gives up after a minute so an unattended screen still presents). Installing closes the running launcher, and the installer reopens what was open.</li>
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
<p>All in <code>launcher/</code>: <code>main.go</code> (modes and the watchdog), <code>app.go</code> (the single running instance), <code>tray.go</code> (the notification-area menu), <code>settings.go</code> (the local windows), <code>update.go</code>, <code>config.go</code>, <code>browser.go</code>, and the platform files. <code>go test ./...</code> runs on both platforms.</p>
