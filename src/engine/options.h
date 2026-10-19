#pragma once

#include <cstdint>
#include <optional>
#include <string>
#include <vector>

namespace rethink::engine {

// Diagrammtypen des Layout-Kerns; weitere folgen gemäß docs/charting-ledger.md.
enum class ChartType { Waterfall, Gantt };

// Ausrichtung eines Säulendiagramms: Richtung, in die die Werte wachsen (Einfügemenü: Pfeile um das Symbol).
//   Up = Säulen nach oben (Standard), Down = an der Waagerechten gespiegelt, Right/Left = Balken nach rechts/links.
enum class Orientation { Up, Down, Right, Left };

// ---- Gantt (Zeitplan) ----

enum class TimeUnit { Day, Week, Month, Quarter, Year };
enum class WeekendMode { Plain, Shaded, Hidden };
enum class BarLabel { None, Range, Start, End };
// Gestaltungsvorlagen; alle Farben sind Designfarben des Folienmasters (Akzente, Text, Hintergrund mit Helligkeit).
enum class GanttDesign { Standard, Modern, Contrast, Minimal };
// Kategorien links: Klammer neben fettem Text, farbiger Block mit weißem Text, nur fetter Text
enum class CategoryStyle { Bracket, Block, Text };

// Aussehen eines Balkens oder Meilensteins (Mini-Toolbar beim Rechtsklick). Farben als Designfarbe ("tx1", "bg1",
// "tx2", "bg2", "accent1" … "accent6"), "#RRGGBB" oder "none"; leer = Vorgabe des Designs.
struct ShapeStyle {
    std::wstring fill, line;
    double lineWeight = 0;  // pt; 0 = Vorgabe (1,5 pt, wenn ein Umriss gesetzt ist)
    bool dashed = false;    // gestrichelter Umriss („Dashed Bar“)
    bool empty() const { return fill.empty() && line.empty() && lineWeight == 0 && !dashed; }
};

// Frei gesetztes Zeitelement (nicht an Datenblattdaten gebunden), Teil von "chart.items".
struct GanttItem {
    enum class Kind { Bar, Pentagon, Chevron, Milestone, Bracket, Shade } kind = Kind::Bar;
    int row = 0;             // Zeile in "table" (1 = Meilensteinzeile unter dem Diagramm); Schattierungen: ohne Bedeutung
    std::wstring start;      // ISO-Datum (Meilenstein: sein Datum)
    std::wstring end;        // ISO-Datum, ausschließlich (wie im Datenblatt: Ende 00:00 Uhr)
    std::wstring text;       // Text im Fünfeck/Chevron bzw. an der Klammer; leer = Datum
    std::optional<std::uint32_t> fill;
    ShapeStyle style;
};

// Aussehen je Datenblattzeile (Balken und Meilenstein der Zeile); wandert mit, wenn Zeilen eingefügt, gelöscht,
// verschoben oder sortiert werden.
struct RowStyle {
    int row = 0;
    ShapeStyle bar, milestone;
};

struct GanttOptions {
    std::wstring start, end;           // sichtbarer Bereich (ISO, beide Tage eingeschlossen); leer = an die Daten angepasst
    std::vector<TimeUnit> scales;      // gezeigte Kalenderskalen, grob → fein; leer = automatisch
    WeekendMode weekends = WeekendMode::Plain;  // wie im Original; „leicht grau“ über das Kontextmenü
    int weekendFirst = 6, weekendLast = 7;  // ISO-Wochentage (1 = Montag … 7 = Sonntag)
    int weekStart = 1;                 // erster Wochentag
    GanttDesign design = GanttDesign::Standard;
    std::optional<bool> separators;    // senkrechte Trennlinien an den Skaleneinheiten; ohne Angabe: Design
    std::optional<double> barHeight;   // Balkenhöhe als Anteil der Zeilenhöhe (0,2 … 0,9); ohne Angabe: Design
    std::wstring barShape, milestoneShape;      // Vorgabeformen (kanonische Namen); leer = Design
    std::optional<bool> category;               // Kategoriespalte; ohne Angabe: sichtbar, wenn es Einträge gibt
    std::optional<CategoryStyle> categoryStyle; // ohne Angabe: Design
    std::optional<bool> categoryColors;         // Balken und Meilensteine in der Farbe ihrer Kategorie; ohne Angabe: Design
    bool links = true;                          // Vorgänger (Spalte 7) als Pfeile verbinden
    BarLabel barLabels = BarLabel::None;
    bool milestoneLabels = true;
    std::optional<bool> milestoneLines;  // gestrichelte Linie über alle Zeilen an den Meilensteinen der Meilensteinzeile; ohne Angabe: Design
    std::wstring dateFormat;           // Format der Datumsbeschriftungen an Elementen; leer = nach Sprache (dd.MM.yyyy bzw. M/d/yyyy)
    std::wstring dayFormat, weekFormat, monthFormat, quarterFormat, yearFormat;  // Skalenformate; leer = Vorgabe
    std::optional<bool> responsible, remark;  // Beschriftungsspalten; ohne Angabe: sichtbar, wenn es Einträge gibt
    bool headlines = true;
    std::optional<bool> rowSeparators;  // waagerechte Trennlinien zwischen den Zeilen; ohne Angabe: Design
    std::vector<int> shadedRows;       // Zeilen ("table"-Index) mit Hintergrund
    bool todayLine = false;            // Heute-Linie (gestrichelt)
    std::wstring today;                // ihr Datum (ISO); leer = referenceDay
    std::vector<GanttItem> items;
    std::vector<RowStyle> rowStyles;   // Aussehen je Datenblattzeile (nur Zeilen mit Angaben)
    double referenceDay = 0;           // heutiges Datum (Tage seit 1970), vom Aufrufer gesetzt: Vorgabebereich ohne Daten
};

// Darstellungsoptionen eines gezeichneten rethink-Diagramms ("chart"-Objekt in .ppttc / RETHINK_MODEL).
struct ChartOptions {
    ChartType type = ChartType::Waterfall;
    Orientation orientation = Orientation::Up;
    bool dataTable = false;          // Datentabelle unter dem Diagramm
    std::vector<int> tableRows;      // Datenblattzeilen (Index in "table"), die nur in der Tabelle stehen
    bool segmentLabels = true;
    bool categoryLabels = true;
    bool connectors = true;
    std::optional<int> decimals;     // Nachkommastellen; sonst automatisch
    double fontPt = 12;
    double gap = 0.5;                // Lücke zwischen Säulen relativ zur Säulenbreite
    std::wstring decimalSeparator = L",";
    std::wstring language = L"de";   // Sprache der Kalenderbeschriftungen ("de", "en")
    GanttOptions gantt;
};

}  // namespace rethink::engine
