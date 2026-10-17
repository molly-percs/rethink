#pragma once

#include "util.h"

#include <string>
#include <vector>

namespace rethink::addin {

// Gezeichnete rethink-Diagramme in PowerPoint: derselbe Layout-Kern wie die Kommandozeile.
// Das Element ist eine Gruppe mit Tag RETHINK_MODEL (Typ, Optionen, Tabelle); jede Form darin heißt
// nach ihrer Rolle ("segment r1 c3", "connector c2" …). Wertänderungen passen die vorhandenen Formen
// an Ort und Stelle an, sodass eigene Formatierungen (Farben, Schrift) erhalten bleiben.

// Fügt auf der aktuellen Folie einen Wasserfall mit Beispieldaten und Datentabelle ein und wählt ihn aus.
// Mit openSheet öffnet sich danach das Datenblatt. Liefert den Elementnamen.
std::wstring insertWaterfall(IDispatch* app, bool openSheet);

// Öffnet das eingebaute Datenblatt des ausgewählten Elements (Fenster im Tabellen-Look neben dem Diagramm).
// Jede Eingabe zeichnet das Diagramm sofort neu; Schließen genügt. Mit show=false unsichtbar (Tests).
void openDatasheet(IDispatch* app, bool show);
// Für Tests: Eingabe in die Zelle (row, column wie "table") des offenen Datenblatts, Anzeige lesen, schließen.
void datasheetEnterCell(IDispatch* app, int row, int column, const std::wstring& text);
void datasheetPasteText(IDispatch* app, int row, int column, const std::wstring& text);
std::wstring datasheetCellText(IDispatch* app, int row, int column);
void closeDatasheet(IDispatch* app);
int openDatasheetCount();

// Alternative: Datenblatt in Excel (Automation: Änderungen sofort, Schließen ohne Rückfrage).
// Mit show=false bleibt Excel unsichtbar (Tests). Liefert den Dateipfad.
std::wstring openExcelDatasheet(IDispatch* app, bool show);

// Arbeitsmappe des Datenblatts zum ausgewählten Element (für Tests), sonst nullptr.
ComPtr<IDispatch> datasheetWorkbook(IDispatch* app);

// Übernimmt geänderte Datenblätter sofort (sonst alle 300 ms). Liefert die Zahl neu gezeichneter Elemente.
int applyDatasheets(IDispatch* app);

// Ändert den Wert eines Segments des ausgewählten Elements um einen Schritt (▲ = +1, ▼ = -1);
// large = zehnfacher Schritt. Ist das Datenblatt offen, wird die Excel-Zelle geändert.
void adjustSegment(IDispatch* app, int row, int column, int direction, bool large);

// Segmente eines Elements für die ▲▼-Griffe (Folienkoordinaten in Punkt).
struct SegmentInfo {
    int row = 0, column = 0;
    std::string flags;  // "pos", "neg", "e"
    double left = 0, top = 0, width = 0, height = 0;
};
std::vector<SegmentInfo> segmentsOf(IDispatch* element);

// Zeichnet das Element neu, wenn seine Größe seit dem letzten Zeichnen geändert wurde. Liefert true bei Neuzeichnen.
bool relayoutIfResized(IDispatch* element);

int watchCount();
void stopDatasheetWatch();

// rethink-Element einer PowerPoint-Auswahl (auch wenn eine Form innerhalb der Gruppe markiert ist), sonst nullptr.
ComPtr<IDispatch> rethinkElementOf(IDispatch* selection);

void commandInsertWaterfall(IDispatch* app);
void commandOpenDatasheet(IDispatch* app);

}  // namespace rethink::addin
