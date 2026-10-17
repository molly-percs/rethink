#pragma once

#include "util.h"

#include <string>

namespace rethink::addin {

// Gezeichnete rethink-Diagramme in PowerPoint: derselbe Layout-Kern wie die Kommandozeile,
// die Grundformen werden über das PowerPoint-Objektmodell gezeichnet und gruppiert.
// Die Gruppe trägt das Tag RETHINK_MODEL (Typ, Optionen, Datentabelle).

// Fügt auf der aktuellen Folie einen Wasserfall mit Beispieldaten und Datentabelle ein und wählt ihn aus.
// Mit openSheet öffnet sich danach das Datenblatt in Excel. Liefert den Elementnamen.
std::wstring insertWaterfall(IDispatch* app, bool openSheet);

// Schreibt das Datenblatt des ausgewählten rethink-Elements als Arbeitsmappe und überwacht sie:
// Jedes Speichern in Excel zeichnet das Element neu. Mit show öffnet sich Excel. Liefert den Dateipfad.
std::wstring openDatasheet(IDispatch* app, bool show);

// Prüft alle überwachten Datenblätter sofort (sonst sekündlich) und zeichnet geänderte Elemente neu.
// Liefert die Zahl neu gezeichneter Elemente.
int applyDatasheets(IDispatch* app);

// Zahl der überwachten Datenblätter (für Tests).
int watchCount();

// Beendet die Überwachung (beim Entladen des Add-ins).
void stopDatasheetWatch();

// rethink-Element einer PowerPoint-Auswahl (auch wenn eine Form innerhalb der Gruppe markiert ist), sonst nullptr.
ComPtr<IDispatch> rethinkElementOf(IDispatch* selection);

// Anzeigename eines Elements für das Auswahl-Etikett, z. B. "rethink · Wasserfall1".
std::wstring elementLabel(IDispatch* element);

void commandInsertWaterfall(IDispatch* app);
void commandOpenDatasheet(IDispatch* app);

}  // namespace rethink::addin
