#pragma once

#include "util.h"

#include <string>
#include <vector>

namespace rethink::addin {

// Abonniert PowerPoint-Ereignisse (EApplication):
//  - Auswahlwechsel: ein ausgewähltes rethink-Element erhält den blauen Leuchtrahmen (overlay.h);
//  - Doppelklick auf ein rethink-Element öffnet sein Datenblatt statt der Gruppenbearbeitung.
void startEvents(IDispatch* app);
void stopEvents();

// Diagnose für Tests: Verbindung, Ereignis-IDs, Anzahl empfangener Ereignisse, Overlay sichtbar.
std::wstring eventStatus();

// Diagnose für Tests: Overlay-Fenster, Elementrechteck und alle Griffe in Bildschirmpixeln.
std::wstring overlayInfo();

// Für Tests: zieht den ▲▼-Griff des Segments (row, column) des ausgewählten Elements über dieselben Schritte wie die
// Maus (pixels: Mausweg je Bewegung, positiv = nach oben) und lässt los. Liefert den eingefrorenen Maßstab.
std::wstring simulateSegmentDrag(int row, int column, const std::vector<int>& pixels);

}  // namespace rethink::addin
