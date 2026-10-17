#pragma once

#include "util.h"

namespace rethink::addin {

// Abonniert PowerPoint-Ereignisse (EApplication):
//  - Auswahlwechsel: ein ausgewähltes rethink-Element erhält den blauen Leuchtrahmen (overlay.h);
//  - Doppelklick auf ein rethink-Element öffnet sein Datenblatt statt der Gruppenbearbeitung.
void startEvents(IDispatch* app);
void stopEvents();

// Diagnose für Tests: Verbindung, Ereignis-IDs, Anzahl empfangener Ereignisse, Overlay sichtbar.
std::wstring eventStatus();

}  // namespace rethink::addin
