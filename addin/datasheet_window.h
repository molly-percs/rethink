#pragma once

#include "ppttc.h"

#include <windows.h>

#include <functional>
#include <string>
#include <vector>

namespace rethink::addin {

// Eingebautes Datenblatt im Tabellenkalkulations-Look: Raster mit Spaltenbuchstaben und Zeilennummern,
// Eingabe direkt in der Zelle, Pfeiltasten/Enter/Tab, Strg+C/V (tabulatorgetrennt, auch aus Excel),
// Strg+Z, Zeilen/Spalten einfügen und löschen, Prozent, Zellfarbe. Jede Änderung meldet sofort die
// ganze Tabelle; Schließen genügt – es gibt nichts zu speichern.
struct DatasheetCallbacks {
    std::function<void(const std::vector<Row>&)> changed;  // nach jeder Änderung
    std::function<void()> openInExcel;                      // Kontextmenü „In Excel bearbeiten“
    std::function<void()> closed;
};

class Datasheet;

// beside: Bildschirmrechteck des Diagramms (das Fenster erscheint daneben). hint: Hinweiszeile unten (leer = Vorgabe).
Datasheet* openDatasheetWindow(HWND owner, const std::wstring& title, std::vector<Row> table, const std::wstring& decimalSeparator,
                               DatasheetCallbacks callbacks, const RECT& beside, bool show, const std::wstring& hint = L"");
void datasheetSetTable(Datasheet* sheet, std::vector<Row> table);  // von außen geändert (z. B. ▲▼), ohne Rückmeldung
void datasheetActivate(Datasheet* sheet);
void datasheetClose(Datasheet* sheet);
HWND datasheetWindow(Datasheet* sheet);

// Für Tests: wie Eingabe + Enter in Zelle (row, column; 0-basiert wie "table") bzw. Anzeige lesen.
void datasheetEnter(Datasheet* sheet, int row, int column, const std::wstring& text);
void datasheetPaste(Datasheet* sheet, int row, int column, const std::wstring& tabSeparated);
std::wstring datasheetText(Datasheet* sheet, int row, int column);

}  // namespace rethink::addin
