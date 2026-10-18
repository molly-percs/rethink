#pragma once

#include <windows.h>

#include <functional>
#include <string>

namespace rethink::addin {

// Kalender eines Zeitplans (Doppelklick auf Skala oder Element): Monatskalender mit Bereichsauswahl (Umschalt+Klick,
// Ziehen über Tage), oben rechts Anfang und Ende in ISO-Schreibweise und die Länge. Jede Auswahl gilt sofort;
// bei einem einzelnen Datum (Meilenstein) schließt sich das Fenster danach. Schließen mit der Schaltfläche, Esc
// oder einem Klick in PowerPoint. Es ist höchstens ein Kalender offen.
struct CalendarCallbacks {
    std::function<void(double first, double last)> changed;  // Tage seit 1970, beide eingeschlossen
};

void openCalendarWindow(HWND owner, const std::wstring& title, double first, double last, bool single, CalendarCallbacks callbacks,
                        const RECT& beside);
void closeCalendarWindow();
bool calendarWindowOpen();

}  // namespace rethink::addin
