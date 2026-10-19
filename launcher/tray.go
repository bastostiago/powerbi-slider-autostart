package main

import (
	"log"
	"strings"

	"fyne.io/systray"
)

// runTray shows the Power Slider icon in the notification area while presenting. Its menu is the
// way out of a kiosk window, which has no other control: press the Windows key to bring the taskbar
// over the presentation, then use the icon. On Linux the icon needs a desktop that shows tray icons
// (StatusNotifierItem: KDE, XFCE, Cinnamon, Ubuntu's GNOME); without one, presenting works the
// same and Alt+F4 three times remains the way out.
func runTray(p *presenter) (err error) {
	if !trayAvailable() {
		log.Printf("no tray on this desktop: presenting without the icon")
		return p.watch()
	}
	tr := trayText(systemLanguage())
	// The watchdog does not wait for the tray: a desktop without a tray host must still present.
	done := make(chan error, 1)
	go func() {
		done <- p.watch()
		systray.Quit()
	}()
	// Whatever goes wrong in the tray library must not take the presentation down with it.
	defer func() {
		if r := recover(); r != nil {
			log.Printf("tray: %v", r)
			err = <-done
		}
	}()
	systray.Run(func() {
		systray.SetIcon(trayIcon)
		systray.SetTooltip(tr.tooltip)
		stop := systray.AddMenuItem(tr.stop, "")
		restart := systray.AddMenuItem(tr.restart, "")
		systray.AddSeparator()
		settings := systray.AddMenuItem(tr.settings, "")
		go func() {
			for {
				select {
				case <-stop.ClickedCh:
					log.Printf("tray: stop")
					p.stop()
				case <-restart.ClickedCh:
					log.Printf("tray: restart")
					p.restart()
				case <-settings.ClickedCh:
					if err := startDetached(); err != nil {
						log.Printf("tray: settings: %v", err)
					}
				}
			}
		}()
	}, func() {})
	// Run also returns when the tray could not start; the presentation goes on until it ends.
	return <-done
}

type trayStrings struct{ tooltip, stop, restart, settings string }

// The site's eight languages, as in ui/i18n.js.
var trayStringsByLang = map[string]trayStrings{
	"en":    {"Power Slider: presentation running", "Stop presentation", "Restart presentation", "Settings"},
	"pt-BR": {"Power Slider: apresentação em andamento", "Parar apresentação", "Reiniciar apresentação", "Configurações"},
	"es":    {"Power Slider: presentación en curso", "Detener presentación", "Reiniciar presentación", "Configuración"},
	"fr":    {"Power Slider : présentation en cours", "Arrêter la présentation", "Redémarrer la présentation", "Paramètres"},
	"de":    {"Power Slider: Präsentation läuft", "Präsentation beenden", "Präsentation neu starten", "Einstellungen"},
	"it":    {"Power Slider: presentazione in corso", "Interrompi presentazione", "Riavvia presentazione", "Impostazioni"},
	"nl":    {"Power Slider: presentatie loopt", "Presentatie stoppen", "Presentatie opnieuw starten", "Instellingen"},
	"ar":    {"Power Slider: العرض قيد التشغيل", "إيقاف العرض", "إعادة تشغيل العرض", "الإعدادات"},
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
