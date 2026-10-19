#pragma once

#include "engine/options.h"
#include "ppttc.h"
#include "util.h"

#include <string>
#include <vector>

namespace rethink::addin {

// Gezeichnete rethink-Diagramme in PowerPoint: derselbe Layout-Kern wie die Kommandozeile.
// Das Element ist eine Gruppe mit Tag RETHINK_MODEL (Typ, Optionen, Tabelle); jede Form darin heißt
// nach ihrer Rolle ("segment r1 c3", "connector c2" …). Wertänderungen passen die vorhandenen Formen
// an Ort und Stelle an, sodass eigene Formatierungen (Farben, Schrift) erhalten bleiben.

// Einfügbare Elemente (Einfügemenü); nicht umgesetzte Typen erscheinen dort ausgegraut.
enum class ElementKind { Waterfall, WaterfallDown, Gantt };

// Fügt auf der aktuellen Folie ein Element mit Beispieldaten ein und wählt es aus. Ausrichtung nur für Säulendiagramme.
// Mit openSheet öffnet sich danach das Datenblatt. Liefert den Elementnamen.
std::wstring insertElement(IDispatch* app, ElementKind kind, engine::Orientation orientation, bool openSheet);
std::wstring insertWaterfall(IDispatch* app, bool openSheet);

// Art eines Elements: "waterfall" oder "gantt" und die Richtung wachsender Werte (Säulendiagramme).
struct ElementInfo {
    bool gantt = false;
    engine::Orientation direction = engine::Orientation::Up;
};
ElementInfo elementInfo(IDispatch* element);

// Öffnet das eingebaute Datenblatt des ausgewählten Elements (Fenster im Tabellen-Look neben dem Diagramm).
// Jede Eingabe zeichnet das Diagramm sofort neu; Schließen genügt. Mit show=false unsichtbar (Tests).
void openDatasheet(IDispatch* app, bool show);
// Für Tests: Eingabe in die Zelle (row, column wie "table") des offenen Datenblatts, Anzeige lesen, schließen.
void datasheetEnterCell(IDispatch* app, int row, int column, const std::wstring& text);
void datasheetMenuCommand(IDispatch* app, int row, const std::wstring& what);  // Tests: Zeile einfügen/löschen, Rückgängig
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

// Ändert den Wert eines Segments des ausgewählten Elements um einen Schritt (▲ = +1, ▼ = -1); die Schrittweite
// folgt der Genauigkeit des eingegebenen Werts (stepDecimals), large = zehnfacher Schritt.
// Ist das Datenblatt offen, wird die Excel-Zelle geändert.
void adjustSegment(IDispatch* app, int row, int column, int direction, bool large);

// Segmente eines Elements für die ▲▼-Griffe (Folienkoordinaten in Punkt).
struct SegmentInfo {
    int row = 0, column = 0;
    std::string flags;  // "pos", "neg", "e"
    double left = 0, top = 0, width = 0, height = 0;
};
std::vector<SegmentInfo> segmentsOf(IDispatch* element);

// Ziehen eines ▲▼-Griffs am ausgewählten Element. Beginn friert Modell, Layout-Rahmen und Maßstab ein und
// öffnet einen Rückgängig-Eintrag für das ganze Ziehen; wirft bei "e" und Nicht-Zahlen.
struct SegmentDragStart {
    double value = 0;          // Wert zu Beginn
    int decimals = 0;          // Schrittweite 10^-decimals (stepDecimals)
    double pointsPerUnit = 0;  // eingefrorener Maßstab; 0, wenn alle Segmente 0 sind
};
SegmentDragStart beginSegmentDrag(IDispatch* app, int row, int column);
// Setzt den (bereits gerundeten) Wert und passt nur geänderte Formen an. Liefert false, wenn sich nichts
// geändert hat; sonst die neue Lage der Segmente (ohne die Formen erneut abzufragen).
bool dragSegmentTo(double value, std::vector<SegmentInfo>& segments);
// Schreibt Modell und Layout an das Element und führt offene Datenblätter (eingebaut, Excel) nach.
void endSegmentDrag();

// Zeichnet das Element neu, wenn seine Größe seit dem letzten Zeichnen geändert wurde. Liefert true bei Neuzeichnen.
bool relayoutIfResized(IDispatch* element);

// ---- Zeitplan (Gantt) ----

// Ziehbare Formen eines Zeitplans (Folienpunkte) und das Band der Kalenderskalen.
struct GanttShape {
    std::string role, flags;  // flags: "bar", "milestone", "bracket", "shade"
    double left = 0, top = 0, width = 0, height = 0;
};
struct GanttGeometry {
    bool valid = false;
    std::vector<GanttShape> shapes;
    double plotLeft = 0, plotRight = 0, scaleTop = 0, scaleBottom = 0;
    double rowsTop = 0, rowH = 0, labelsLeft = 0;  // Aktivitätszeilen (Zeile 2 ab rowsTop) und linker Rand der Beschriftungen
    int rowCount = 0;
};
GanttGeometry ganttGeometryOf(IDispatch* element);

// Ziehen am ausgewählten Zeitplan. grip: 0 = Anfang, 1 = Ende, 2 = verschieben; role "range" = sichtbarer Bereich
// (grip 0/1). pointsPerPixel: Folienpunkte je Bildschirmpixel (Rasterung: nicht feiner als ~3 Pixel).
// Ein Rückgängig-Eintrag für das ganze Ziehen. Liefert die Kurzinfo zu Beginn.
std::wstring beginGanttDrag(IDispatch* app, const std::string& role, int grip, double pointsPerPixel);
// Mausweg dx (Folienpunkte). Formen werden an Ort und Stelle angepasst (Bereich: erst beim Loslassen).
// Liefert true, wenn sich die Daten geändert haben; geometry/tooltip enthalten dann den neuen Stand.
bool ganttDragTo(double dx, GanttGeometry& geometry, std::wstring& tooltip);
// Übernimmt (oder verwirft bei cancel) das Ziehen und führt offene Datenblätter nach.
void endGanttDrag(bool cancel);

// Was liegt an einem Punkt (Folienpunkte) eines Zeitplans: Zeile, Datum, Form, Skalenband.
struct GanttHit {
    bool valid = false;
    int row = -1;         // "table"-Index (1 = Meilensteinzeile), -1 = keine Zeile
    double date = 0;      // Tage seit 1970
    bool hasDate = false;
    std::string role;     // ziehbare Form unter dem Punkt (leer = keine)
    std::string label;    // Datumsbeschriftung unter dem Punkt ("bar r2 label", "milestone label r3" …), wenn keine Form
    bool inScale = false; // im Band der Kalenderskalen
};
GanttHit ganttHitAt(IDispatch* element, double x, double y);

// Befehle am ausgewählten Zeitplan; hit: Mausposition beim Öffnen des Menüs. Ein Befehl ist "schlüssel=wert" oder ein
// Aktionsname, z. B. "design=modern", "barHeight=0.75", "weekends=shaded", "separators=0", "rowSeparators=1",
// "scale:day=1", "scales=auto", "format:month=MMM", "dateFormat=dd.MM.yy", "barLabels=range", "column:category=1",
// "categoryStyle=block", "categoryColors=1", "links=1", "link=Analyse" (Vorgänger der Zeile umschalten), "rowShape=chevron",
// "rowMarker=star", "barShape=rounded", "milestoneShape=circle", "rowShade=1", "insertRow", "duplicateRow", "deleteRow",
// "rowUp", "rowDown", "moveRow=5" (Zeile an Index 5, übernimmt dort die Kategorie), "new=bar|milestone|bracket|shade|pentagon|chevron", "deleteItem", "fit", "calendar", "datasheet",
// "sortCategories" (Dialog) bzw. "sortCategories=Bau|Planung". Aussehen der angeklickten Form (Minileiste): "fill=accent2",
// "line=#C00000", "fill=none", "fill=" (Vorgabe), "lineWeight=0.75", "dashed=1", "barStyle=dashed|solid". Die IDs des PowerPoint-Kontextmenüs (rtkG_…) gehen auch.
void ganttCommand(IDispatch* app, const std::wstring& command, const GanttHit& hit);
// Umschalter des PowerPoint-Kontextmenüs (z. B. "rtkT_separators"): Zustand lesen bzw. setzen.
bool ganttPressed(IDispatch* app, const std::wstring& id, const GanttHit& hit);
void ganttToggle(IDispatch* app, const std::wstring& id, bool pressed, const GanttHit& hit);

// Zustand des ausgewählten Zeitplans für das eigene Kontextmenü (Häkchen, Auswahllisten).
struct GanttState {
    engine::GanttOptions options;          // gespeicherte Optionen
    std::vector<engine::TimeUnit> scales;  // gezeigte Skalen, grob → fein
    double barHeight = 0.55;               // wirksam (Optionen oder Design)
    std::wstring barShape, marker;
    engine::CategoryStyle categoryStyle = engine::CategoryStyle::Bracket;
    bool categoryColors = false, separators = true, rowSeparators = false, milestoneLines = false;
    bool categoryShown = false, responsibleShown = false, remarkShown = false;
    std::vector<std::wstring> activities;  // Beschriftung je Aktivitätszeile (Zeile 2 = Index 0)
    std::vector<std::wstring> categories;
    std::string targetKind;                // Form unter der Maus: "bar", "milestone", "bracket", "shade" oder leer
    bool targetIsItem = false;             // frei gesetztes Element (sonst Datenblatt)
    std::wstring rowShape, rowMarker;      // Spalten 8/9 der Zeile unter der Maus (kanonisch, leer = Vorgabe)
    std::vector<std::wstring> rowLinks;    // Vorgänger der Zeile
    bool rowShaded = false;
    engine::ShapeStyle targetStyle;        // gespeichertes Aussehen der Form unter der Maus (leer = Vorgabe)
    long targetFill = -1, targetLine = -1; // wie gezeichnet (RGB wie PowerPoint, -1 = keine)
    double targetWeight = 0;
    std::vector<std::pair<std::wstring, long>> theme;  // Designfarben der Folie: tx1, bg1, tx2, bg2, accent1 … accent6
    std::wstring defaultFill;              // Vorgabefarbe der Form unter der Maus (Designfarbe des Masters, z. B. "accent2")
    double today = 0;
    std::wstring language;
};
GanttState ganttState(IDispatch* app, const GanttHit& hit);

// Kalender (Doppelklick auf Skala oder Element, Kontextmenü): role leer = sichtbarer Bereich.
void openGanttCalendar(IDispatch* app, const std::string& role);
// Für Tests: Bereich bzw. Daten eines Elements setzen wie im Kalender (ISO, beide Tage eingeschlossen).
void setGanttDates(IDispatch* app, const std::string& role, const std::wstring& first, const std::wstring& last);
// Für Tests: Ziele des ausgewählten Zeitplans mit ihren Daten ("bar r2=2026-11-02..2026-11-20; …").
std::wstring ganttTargets(IDispatch* app);

// Form innerhalb des Elements auswählen (Klick auf einen Balken: wie PowerPoints zweiter Klick in eine Gruppe).
void selectInnerShape(IDispatch* element, const std::string& role);

int watchCount();
void stopDatasheetWatch();

// rethink-Element einer PowerPoint-Auswahl (auch wenn eine Form innerhalb der Gruppe markiert ist), sonst nullptr.
ComPtr<IDispatch> rethinkElementOf(IDispatch* selection);

void commandInsertWaterfall(IDispatch* app);
void commandInsertElement(IDispatch* app, ElementKind kind, engine::Orientation orientation);
void commandOpenDatasheet(IDispatch* app);

}  // namespace rethink::addin
