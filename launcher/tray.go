package main

import (
	"log"
	"strings"
	"sync"

	"fyne.io/systray"
)

// trayMenu is the Power Slider icon in the notification area, shown for as long as the launcher
// runs. Its menu follows the state: Stop and Restart while presenting, Start when not, Settings
// always. In a kiosk window the Windows key brings the taskbar over the presentation. On Linux it
// needs a desktop that shows tray icons (StatusNotifierItem: KDE, XFCE, Cinnamon, Ubuntu's GNOME);
// elsewhere the launcher runs without it (see trayAvailable).
type trayMenu struct {
	start, stop, restart, settings *systray.MenuItem
}

var (
	trayMu      sync.Mutex
	trayRunning bool
)

// runTray blocks until the launcher exits.
func runTray(a *app) {
	tr := trayText(systemLanguage())
	// Whatever goes wrong in the tray library must not take the launcher down with it.
	defer func() {
		if r := recover(); r != nil {
			log.Printf("tray: %v", r)
			<-a.quit
		}
	}()
	systray.Run(func() {
		systray.SetIcon(trayIcon)
		systray.SetTooltip(tr.tooltip)
		m := &trayMenu{
			start:   systray.AddMenuItem(tr.start, ""),
			stop:    systray.AddMenuItem(tr.stop, ""),
			restart: systray.AddMenuItem(tr.restart, ""),
		}
		systray.AddSeparator()
		m.settings = systray.AddMenuItem(tr.settings, "")
		trayMu.Lock()
		trayRunning = true
		a.tray = m
		trayMu.Unlock()
		a.refreshTray()
		go func() {
			for {
				var err error
				select {
				case <-m.start.ClickedCh:
					err = a.startPresentation()
				case <-m.stop.ClickedCh:
					a.stopPresentation()
				case <-m.restart.ClickedCh:
					a.restartPresentation()
				case <-m.settings.ClickedCh:
					err = a.showSettings()
				case <-a.quit:
					return
				}
				if err != nil {
					log.Printf("tray: %v", err)
				}
			}
		}()
	}, func() {})
}

// refreshTray shows the menu items that apply right now.
func (a *app) refreshTray() {
	trayMu.Lock()
	m := a.tray
	trayMu.Unlock()
	if m == nil {
		return
	}
	if a.presenting() {
		m.start.Hide()
		m.stop.Show()
		m.restart.Show()
	} else {
		m.start.Show()
		m.stop.Hide()
		m.restart.Hide()
	}
}

func quitTray() {
	trayMu.Lock()
	running := trayRunning
	trayMu.Unlock()
	if running {
		systray.Quit()
	}
}

type trayStrings struct{ tooltip, start, stop, restart, settings string }

// The site's eight languages, as in ui/i18n.js.
var trayStringsByLang = map[string]trayStrings{
	"en":    {"Power Slider Launcher", "Start presentation", "Stop presentation", "Restart presentation", "Settings"},
	"pt-BR": {"Power Slider Launcher", "Iniciar apresentação", "Parar apresentação", "Reiniciar apresentação", "Configurações"},
	"es":    {"Power Slider Launcher", "Iniciar presentación", "Detener presentación", "Reiniciar presentación", "Configuración"},
	"fr":    {"Power Slider Launcher", "Démarrer la présentation", "Arrêter la présentation", "Redémarrer la présentation", "Paramètres"},
	"de":    {"Power Slider Launcher", "Präsentation starten", "Präsentation beenden", "Präsentation neu starten", "Einstellungen"},
	"it":    {"Power Slider Launcher", "Avvia presentazione", "Interrompi presentazione", "Riavvia presentazione", "Impostazioni"},
	"nl":    {"Power Slider Launcher", "Presentatie starten", "Presentatie stoppen", "Presentatie opnieuw starten", "Instellingen"},
	"ar":    {"Power Slider Launcher", "بدء العرض", "إيقاف العرض", "إعادة تشغيل العرض", "الإعدادات"},
}

// trayText picks the strings for a locale such as "pt-BR", "pt_PT" or "de-AT", falling back to the
// language alone and then to English.
func trayText(locale string) trayStrings {
	locale = normalizeLocale(locale)
	for key, s := range trayStringsByLang {
		if strings.EqualFold(key, locale) {
			return s
		}
	}
	base, _, _ := strings.Cut(locale, "-")
	for key, s := range trayStringsByLang {
		if strings.EqualFold(strings.SplitN(key, "-", 2)[0], base) {
			return s
		}
	}
	return trayStringsByLang["en"]
}

// normalizeLocale turns "pt_BR.UTF-8" or "de_AT@euro" into "pt-BR" / "de-AT".
func normalizeLocale(locale string) string {
	if cut := strings.IndexAny(locale, ".@"); cut >= 0 {
		locale = locale[:cut]
	}
	return strings.ReplaceAll(locale, "_", "-")
}
