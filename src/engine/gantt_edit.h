#pragma once

// Bearbeiten eines Zeitplans (reine Rechnung): was hinter einer gezeichneten Form steht, Ziehen in Tagen mit
// Rasterung, Datumsbereich, Zeilen einfügen/löschen/verschieben. Das Add-in ruft nur diese Funktionen auf.

#include "chart.h"

#include <string>

namespace rethink::engine {

// Was hinter einer Form steht: Datenblattzellen ("bar r3": Start/Ende, "milestone r3": Meilenstein) oder ein frei
// gesetztes Element ("item 2" in chart.items).
struct GanttTarget {
    enum class Kind { None, Bar, Milestone, Bracket, Shade } kind = Kind::None;
    int row = -1;   // Zeile in "table" (für Datenblatt-Daten), sonst die Zeile des Elements
    int item = -1;  // Index in options.gantt.items, -1 = Datenblatt
    double start = 0, end = 0;  // Meilenstein: start
};
GanttTarget ganttTarget(const DataItem& item, const ChartOptions& options, const std::string& role);
// Schreibt neue Daten zurück (Datenblattzellen behalten ihre Farbe).
void setGanttDates(DataItem& item, ChartOptions& options, const GanttTarget& target, double start, double end);

enum class GanttGrip { Start, End, Move };

// Neue Daten nach einem Zug um dx EMU auf der (eingefrorenen) Achse, gerastet auf `unit`.
// Ende bleibt nach dem Start (mindestens eine Einheit bzw. ein Tag), Verschieben erhält die Dauer.
std::pair<double, double> ganttDragDates(const GanttAxis& axis, const GanttTarget& target, GanttGrip grip, double dx, TimeUnit unit);

// Neuer Anfang (startSide) bzw. neues Ende des sichtbaren Bereichs nach einem Zug an der Skala, gerastet auf `unit`;
// der Bereich bleibt mindestens eine Einheit lang. Liefert [t0, t1).
std::pair<double, double> ganttRangeDrag(const GanttAxis& axis, bool startSide, double dx, TimeUnit unit);

// Text der Kurzinfo beim Ziehen, z. B. "Mi 05.11.2026 · KW 45".
std::wstring ganttTooltip(double t, const std::wstring& language);

// Sichtbaren Bereich setzen (beide Tage eingeschlossen, ISO in options.gantt.start/end).
void setGanttRange(ChartOptions& options, double t0, double t1);
// "Skala an Daten anpassen": Bereich = früheste bis späteste Angabe; false, wenn es keine Daten gibt.
bool fitGanttRange(const DataItem& item, ChartOptions& options);

// Zeilen (Index in "table", ab gantt::kFirstActivityRow). Elemente in betroffenen Zeilen wandern mit.
int insertGanttRow(DataItem& item, ChartOptions& options, int before, const std::wstring& label);  // liefert die neue Zeile
void duplicateGanttRow(DataItem& item, ChartOptions& options, int row, bool above);
void deleteGanttRow(DataItem& item, ChartOptions& options, int row);
bool moveGanttRow(DataItem& item, ChartOptions& options, int row, int delta);  // false am Rand

// Neues Element an einem Datum in einer Zeile (Kontextmenü an der Mausposition). Balken und Meilensteine landen
// im Datenblatt, wenn die Zeile dort noch keine hat; sonst als frei gesetztes Element. Dauer neuer Balken: `days`.
void addGanttItem(DataItem& item, ChartOptions& options, GanttItem::Kind kind, int row, double date, double days);
// Löscht das Element hinter einer Rolle (Datenblattzellen werden geleert).
void deleteGanttTarget(DataItem& item, ChartOptions& options, const GanttTarget& target);

// Text in eine Datenblattzelle schreiben (leer = Zelle leeren, Farbe bleibt), z. B. Form (Spalte 8) oder Marker (9).
void setGanttText(DataItem& item, int row, int column, const std::wstring& text);

// Vorgänger (Spalte 7) umschalten: ist `predecessor` eingetragen, wird er entfernt, sonst angehängt. Liefert den neuen Zustand.
bool toggleGanttLink(DataItem& item, int row, const std::wstring& predecessor);

// Aktivitätszeilen nach Kategorie sortieren: zuerst die Kategorien in `order`, dann übrige in Reihenfolge ihres Auftretens,
// Zeilen ohne Kategorie zuletzt; innerhalb einer Kategorie bleibt die Reihenfolge. Frei gesetzte Elemente wandern mit.
void sortGanttByCategory(DataItem& item, ChartOptions& options, const std::vector<std::wstring>& order);

}  // namespace rethink::engine
