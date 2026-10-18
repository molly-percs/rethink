#pragma once

// Layout-Kern: reine Berechnung ohne COM und ohne XML.
// Ergebnis ist eine Liste von Grundformen in EMU, die die Ausgaben (OOXML, später COM) nur umsetzen.

#include "../ppttc.h"
#include "options.h"

#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

namespace rethink::engine {

using Emu = std::int64_t;
inline constexpr Emu kEmuPerPt = 12700;

struct Box {
    Emu x = 0, y = 0, w = 0, h = 0;
};

struct Color {
    enum class Kind { None, Scheme, Rgb } kind = Kind::None;
    std::wstring scheme;  // "accent1", "tx1", "bg1" … – Designfarben folgen dem Folienmaster
    std::uint32_t rgb = 0;
    double brightness = 0;  // Schemafarben: -1 (dunkler) … +1 (heller), wie „Heller 75 %“ in PowerPoint

    static Color none() { return {}; }
    static Color schemeColor(std::wstring name, double brightness = 0) {
        Color c;
        c.kind = Kind::Scheme;
        c.scheme = std::move(name);
        c.brightness = brightness;
        return c;
    }
    static Color rgbColor(std::uint32_t value) {
        Color c;
        c.kind = Kind::Rgb;
        c.rgb = value;
        return c;
    }
};

enum class Align { Left, Center, Right };

struct Prim {
    // Shape: vordefinierte Form `shape` (OOXML-Name: "diamond", "triangle", "homePlate", "chevron"), optional mit Text
    enum class Kind { Rect, Line, Text, Shape } kind = Kind::Rect;
    Emu x = 0, y = 0, w = 0, h = 0;  // Linie: von (x, y) nach (x + w, y + h)
    Color fill, stroke, textColor;
    double transparency = 0;  // Füllung: 1 = unsichtbar, aber anklickbar (Fläche hinter dem Diagramm)
    double strokePt = 0.75;
    bool dashed = false;
    std::wstring text;
    double fontPt = 12;
    Align align = Align::Center;
    bool bold = false;
    std::string shape;    // nur Kind::Shape
    double rotation = 0;  // Grad im Uhrzeigersinn (Kind::Shape)
    std::string role;   // eindeutiger Rollenname, z. B. "segment r1 c3"; Formname im Element
    std::string flags;  // Segmente: "pos", "neg" oder "e" (berechnet); Gantt: "bar", "milestone" … – für Griffe im Add-in
    double value = 0;   // Segmente: dargestellter Wert (für den Maßstab beim Ziehen)
};

// Gleich gezeichnet? (Geometrie, Text, Farben, Kennung) – für Aktualisierungen nur geänderter Formen.
bool samePrim(const Prim& a, const Prim& b);

// Zeitachse eines Zeitplans: sichtbarer Bereich [t0, t1) in Tagen auf die Breite [x0, x0 + w] in EMU.
// Ausgeblendete Wochenenden haben keine Breite.
struct GanttAxis {
    bool valid = false;
    double t0 = 0, t1 = 1;
    Emu x0 = 0, w = 1;
    bool hideWeekends = false;
    int weekendFirst = 6, weekendLast = 7, weekStart = 1;
    TimeUnit finest = TimeUnit::Day;  // feinste gezeigte Skala
    bool dayScale = false;            // Tagesskala sichtbar: Meilensteine in der Tagesmitte

    double xOf(double t) const;       // EMU (nicht gerundet), außerhalb des Bereichs linear fortgesetzt
    double dateAt(double x) const;    // Umkehrung von xOf
    double emuPerDay() const;
    // Kleinste sinnvolle Rasterung beim Ziehen: Tag, sonst Woche oder Monat, wenn ein Tag schmaler als minEmu ist.
    TimeUnit snapUnit(double minEmu) const;
};

// Zeitplan: senkrechte Lage einer Zeile ("table"-Index; 1 = Meilensteinzeile) – für Mausposition → Zeile im Add-in.
struct RowBand {
    int row = 0;
    Emu y = 0, h = 0;
};

struct Drawing {
    Box bounds;
    std::vector<Prim> prims;
    std::vector<std::string> warnings;
    Orientation valueDirection = Orientation::Up;  // Säulendiagramme: Richtung wachsender Werte (für ▲▼-Griffe)
    GanttAxis axis;                                // Zeitplan: Zeitachse (für Ziehen in Tagen)
    std::vector<RowBand> rows;                     // Zeitplan: Zeilen
    Emu scaleTop = 0, scaleBottom = 0;             // Zeitplan: Band der Kalenderskalen
    std::vector<TimeUnit> scales;                  // Zeitplan: gezeigte Skalen (grob → fein)
};

// ---- Wasserfall ----------------------------------------------------------

struct WaterfallSegment {
    int row = 0;                // Index in "table"
    double value = 0;           // eingegebener oder berechneter Wert
    double lo = 0, hi = 0;      // Ausdehnung in Dateneinheiten
    bool computed = false;      // aus "e" berechnet
    bool percent = false;       // Eingabe war "percentage"
    std::optional<std::uint32_t> fill;
};

struct WaterfallColumn {
    std::vector<WaterfallSegment> segments;
    bool anchored = false;      // Säule mit "e": steht auf der Basislinie (Zwischensumme)
    double levelBefore = 0;
    double levelAfter = 0;      // Höhe des Verbinders zur nächsten Säule
};

// Wasserfall-Arithmetik nach dem dokumentierten Verhalten:
//  - Säulen ohne "e" schweben ab dem laufenden Stand und verschieben ihn um ihre Summe.
//  - Ein "e" ergänzt seine Säule bis zur Basislinie (Zwischensumme); der Stand bleibt.
//  - Ein "e" in der ersten Säule wird rückwärts berechnet (der Stand endet bei 0).
//  - Höchstens ein "e" pro Säule; gemischte Vorzeichen stapeln getrennt nach oben und unten.
std::vector<WaterfallColumn> computeWaterfall(const DataItem& item, const ChartOptions& options,
                                              std::vector<std::string>* warnings = nullptr);

Drawing layoutChart(const DataItem& item, const ChartOptions& options, Box bounds);

// Nachkommastellen der Beschriftungen (fest aus den Optionen oder automatisch 0–2);
// das Add-in nutzt sie als Schrittweite der ▲▼-Griffe (10^-Stellen).
int chartDecimals(const DataItem& item, const ChartOptions& options);

// Schrittweite beim Ändern eines einzelnen Werts (▲▼, Ziehen): so viele Nachkommastellen, wie der Wert selbst
// eingegeben wurde (12 → 1er-Schritte, 12,5 → 0,1); eine feste "decimals"-Option hat Vorrang. Höchstens 4.
int stepDecimals(const Cell& cell, const ChartOptions& options);

// Wert beim Ziehen: Startwert + Mausweg (in Pixeln, positiv = Richtung wachsender Werte) / Pixel je Einheit,
// gerundet auf die Schrittweite. Der Maßstab wird zu Beginn eingefroren.
double dragValue(double startValue, double pixels, double pixelsPerUnit, int decimals);

// Maßstab einer Zeichnung in EMU je Werteinheit, abgelesen am Segment mit dem größten Betrag
// (0, wenn alle Segmente 0 sind). Grundlage des eingefrorenen Maßstabs beim Ziehen.
double emuPerUnit(const Drawing& drawing);

// Zahl mit festen Nachkommastellen und gewähltem Dezimaltrennzeichen.
std::wstring formatValue(double value, int decimals, const std::wstring& decimalSeparator, bool percent);

// Ist die Zelle ein "e" (berechnetes Segment)?
bool isComputedCell(const Cell& cell);

// ---- Zeitplan (Gantt) ----
//
// Datenblatt wie im Original: Zeile 0 = Spaltenüberschriften (Aktivität, Start, Ende, Meilenstein, Verantwortlich,
// Bemerkung – die Überschriften der Beschriftungsspalten erscheinen im Diagramm), Zeile 1 = Meilensteinzeile unter dem
// Diagramm, ab Zeile 2 je Zeile eine Aktivität. Spalten: 0 Aktivität, 1 Start, 2 Ende, 3 Meilenstein, 4 Verantwortlich,
// 5 Bemerkung. Ein Datum ohne Uhrzeit ist 00:00 Uhr – ein Balken, der am 20. endet, schließt den 20. nicht ein.

// Weitere Spalten (rethink): 6 Kategorie (Zeilen gruppiert links neben den Aktivitäten), 7 Vorgänger (Aktivitätsnamen,
// mit ; getrennt – als Pfeil verbunden), 8 Form des Balkens (Balken, Abgerundet, Fünfeck, Chevron), 9 Form des
// Meilensteins (Raute, Dreieck, Kreis, Quadrat, Stern). Leere Formzellen nehmen die Vorgabe des Designs.
namespace gantt {
inline constexpr int kHeaderRow = 0, kMilestoneRow = 1, kFirstActivityRow = 2;
inline constexpr int kLabelCol = 0, kStartCol = 1, kEndCol = 2, kMilestoneCol = 3, kResponsibleCol = 4, kRemarkCol = 5;
inline constexpr int kCategoryCol = 6, kPredecessorCol = 7, kShapeCol = 8, kMarkerCol = 9;
inline constexpr int kColumns = 10;
}  // namespace gantt

// Wirksame Gestaltung: Angaben der Optionen, sonst die Vorgaben des Designs.
struct GanttLook {
    double barHeight = 0.55;
    std::wstring barShape, marker;  // kanonische Namen
    CategoryStyle categoryStyle = CategoryStyle::Bracket;
    bool categoryColors = false, separators = true, rowSeparators = false;
};
GanttLook ganttLook(const GanttOptions& options);

// Kategorien in der Reihenfolge ihres ersten Auftretens (leere ausgenommen).
std::vector<std::wstring> ganttCategories(const DataItem& item);
// Formnamen der Spalten 8/9 (deutsch oder englisch) → kanonisch ("bar", "rounded", "pentagon", "chevron" bzw.
// "diamond", "triangle", "circle", "square", "star"); leer, wenn unbekannt.
std::wstring canonicalBarShape(const std::wstring& name);
std::wstring canonicalMarker(const std::wstring& name);

// Datum einer Zelle (Datum oder ISO-Text), sonst nullopt.
std::optional<double> cellDate(const Cell& cell);

// Frühester Anfang und spätestes Ende aller Balken, Meilensteine, Klammern und Schattierungen (Tage), sonst nullopt.
std::optional<std::pair<double, double>> ganttDataRange(const DataItem& item, const ChartOptions& options);

// Sichtbarer Bereich [t0, t1): "start"/"end" der Optionen, sonst die Daten auf ganze Monate erweitert,
// ohne Daten der Monat von options.gantt.referenceDay und die zwei folgenden.
std::pair<double, double> ganttRange(const DataItem& item, const ChartOptions& options);

// Zeitachse, wie layoutGantt sie für diesen Rahmen verwendet (für das Ziehen im Add-in).
GanttAxis ganttAxis(const DataItem& item, const ChartOptions& options, Box bounds);

// Gezeigte Skalen (grob → fein): die Optionen oder automatisch nach Breite je Einheit und Schriftgröße.
std::vector<TimeUnit> ganttScales(double t0, double t1, Emu plotWidth, const ChartOptions& options);

}  // namespace rethink::engine
