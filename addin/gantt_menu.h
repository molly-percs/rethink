#pragma once

#include "elements.h"

#include <windows.h>

#include <string>

namespace rethink::addin {

// Eigenes Kontextmenü des Zeitplans (Rechtsklick ins ausgewählte Diagramm, wie im Original): Einträge je nach Stelle
// (Balken, Meilenstein, Zeile, Skala) und für das ganze Diagramm (Skalen und ihr Format, Linien, Wochenenden,
// Beschriftung, Spalten, Balken, Kategorien, Verbindungen, Design). Führt die Auswahl über ganttCommand aus. Darüber
// erscheint die Minileiste für die Stelle (siehe unten).
void showGanttMenu(IDispatch* app, HWND owner, POINT screen, const GanttHit& hit);

// Für Tests: das Menü an einer Stelle als Text, je Eintrag "Pfad > Eintrag [x]=Befehl", durch Zeilenumbrüche getrennt.
std::wstring ganttMenuText(IDispatch* app, const GanttHit& hit);

// Minileiste über dem Menü (Form, Füllung, Linie, Linienstärke; Datum; Skalen) – erscheint mit showGanttMenu.
// Fenster der offenen Leiste (nullptr = keine), damit der Maus-Hook Klicks darauf durchlässt.
HWND ganttToolbarWindow();
// Für Tests: die Leiste an einer Stelle, je Schaltfläche "[Beschriftung #RRGGBB]" und darunter ihr Menü wie ganttMenuText.
std::wstring ganttToolbarText(IDispatch* app, const GanttHit& hit);

}  // namespace rethink::addin
