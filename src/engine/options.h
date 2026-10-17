#pragma once

#include <optional>
#include <string>
#include <vector>

namespace rethink::engine {

// Diagrammtypen des Layout-Kerns; weitere folgen gemäß docs/charting-ledger.md.
enum class ChartType { Waterfall };

// Darstellungsoptionen eines gezeichneten rethink-Diagramms ("chart"-Objekt in .ppttc / RETHINK_MODEL).
struct ChartOptions {
    ChartType type = ChartType::Waterfall;
    bool dataTable = false;          // Datentabelle unter dem Diagramm
    std::vector<int> tableRows;      // Datenblattzeilen (Index in "table"), die nur in der Tabelle stehen
    bool segmentLabels = true;
    bool categoryLabels = true;
    bool connectors = true;
    std::optional<int> decimals;     // Nachkommastellen; sonst automatisch
    double fontPt = 12;
    double gap = 0.5;                // Lücke zwischen Säulen relativ zur Säulenbreite
    std::wstring decimalSeparator = L",";
};

}  // namespace rethink::engine
