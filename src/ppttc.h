#pragma once

#include <cstdint>
#include <filesystem>
#include <optional>
#include <string>
#include <vector>

namespace rethink {

// Modell einer .ppttc-Datei (JSON-Format der think-cell-Automatisierung).

enum class CellKind { Empty, String, Number, Percentage, Date };

struct Cell {
    CellKind kind = CellKind::Empty;
    std::wstring text;                  // String, Date
    double number = 0;                  // Number, Percentage (Prozentpunkte, 48 = 48 %)
    std::optional<std::uint32_t> fill;  // 0xRRGGBB
};

using Row = std::vector<Cell>;

struct DataItem {
    std::wstring name;
    std::vector<Row> table;
    // rethink-Erweiterung: optionales "chart"-Objekt (gezeichnetes Diagramm), als JSON-Text.
    // think-cell ignoriert unbekannte Schlüssel; leer = kein gezeichnetes Diagramm angefordert.
    std::wstring chartJson;
};

struct SlideSpec {
    std::wstring templateRef;  // Pfad (relativ zur .ppttc-Datei) oder http(s)-URL
    std::vector<DataItem> data;
};

// Liest und validiert die Datei gemäß dem dokumentierten Schema; Fehler nennen den JSON-Pfad.
std::vector<SlideSpec> parsePpttc(const std::filesystem::path& file);

// Text, wie er in einem Textfeld erscheint.
std::wstring cellText(const Cell& cell);

// ---- Modell gezeichneter Diagramme (RETHINK_MODEL) ----

// {"name": …, "chart": {…}, "table": [[…]]} lesen bzw. schreiben – dieselbe Form wie ein .ppttc-Dateneintrag.
DataItem parseDataItemJson(const std::wstring& json);
std::wstring serializeDataItem(const DataItem& item);

// Schlüssel aus `override_` überschreiben die von `stored` (flach, je Schlüssel).
std::wstring mergeChartJson(const std::wstring& stored, const std::wstring& override_);
// Farbe einer Stilangabe: leer, Designfarbe (tx1, bg1, tx2, bg2, accent1 … accent6), "#rrggbb" oder "none"
bool isStyleColor(const std::wstring& value);

}  // namespace rethink

#include "engine/options.h"

namespace rethink {

// "chart"-Objekt in Optionen übersetzen; unbekannte Typen führen zu einem Fehler.
engine::ChartOptions parseChartOptions(const std::wstring& json);

// Die im Add-in änderbaren Zeitplan-Schlüssel (Bereich, Skalen, Anzeige, Elemente) als JSON-Objekt –
// mit mergeChartJson über das gespeicherte "chart" gelegt.
std::wstring ganttEditsJson(const engine::GanttOptions& gantt);

}  // namespace rethink
