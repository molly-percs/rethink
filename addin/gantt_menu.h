#pragma once

#include "elements.h"

#include <windows.h>

#include <string>

namespace rethink::addin {

// Eigenes Kontextmenü des Zeitplans (Rechtsklick ins ausgewählte Diagramm, wie im Original): Einträge je nach Stelle
// (Balken, Meilenstein, Zeile, Skala) und für das ganze Diagramm (Skalen und ihr Format, Linien, Wochenenden,
// Beschriftung, Spalten, Balken, Kategorien, Verbindungen, Design). Führt die Auswahl über ganttCommand aus.
void showGanttMenu(IDispatch* app, HWND owner, POINT screen, const GanttHit& hit);

// Für Tests: das Menü an einer Stelle als Text, je Eintrag "Pfad > Eintrag [x]=Befehl", durch Zeilenumbrüche getrennt.
std::wstring ganttMenuText(IDispatch* app, const GanttHit& hit);

}  // namespace rethink::addin
