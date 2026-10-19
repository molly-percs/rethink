#pragma once

#include <windows.h>

#include <functional>
#include <string>

namespace rethink::addin {

// Kalender eines Zeitplans (Doppelklick auf Skala oder Element): Jahresraster – je Zeile ein Monat, die Tage nach
// Wochentagen ausgerichtet (Montag zuerst), Wochenendspalten grau, heute rot. Bereich wählen durch Ziehen über die Tage,
// Anfang oder Ende des Bereichs greifen und ziehen, Umschalt+Klick setzt den näheren Rand. Unten die Zusammenfassung
// („01.06.2026 bis 30.06.2026 · 4 Wochen 2 Tage“) und der Tag unter der Maus mit Kalenderwoche. Oben Monat zurück/vor,
// Monat und Jahr der obersten Zeile; das Mausrad blättert. Jede Auswahl gilt sofort; bei einem einzelnen Datum
// (Meilenstein) schließt sich das Fenster danach. Schließen mit der Schaltfläche, Esc oder einem Klick in PowerPoint.
// Es ist höchstens ein Kalender offen.
struct CalendarCallbacks {
    std::function<void(double first, double last)> changed;  // Tage seit 1970, beide eingeschlossen
};

void openCalendarWindow(HWND owner, const std::wstring& title, double first, double last, bool single, CalendarCallbacks callbacks,
                        const RECT& beside);
void closeCalendarWindow();
bool calendarWindowOpen();
// Für Tests: Zusammenfassung des offenen Kalenders (leer = zu) bzw. eine Auswahl wie mit der Maus (Tage seit 1970).
std::wstring calendarSummary();
void calendarPick(double first, double last);

}  // namespace rethink::addin
