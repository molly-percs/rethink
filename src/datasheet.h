#pragma once

#include "ppttc.h"

#include <cstdint>
#include <filesystem>
#include <vector>

namespace rethink {

// Datenblatt eines gezeichneten Elements als Excel-Arbeitsmappe (Blatt "Datenblatt"):
// Zeile/Spalte entsprechen genau "table" (Zeile 1 = Kategorien, Spalte A = Reihennamen).
// Prozentwerte erhalten das Format 0 %, Datumswerte ein Datumsformat, "fill" wird zur Zellfarbe.
std::vector<std::uint8_t> datasheetFromTable(const std::vector<Row>& table);

// Umkehrung: liest eine (in Excel bearbeitete und gespeicherte) Arbeitsmappe in eine Tabelle.
std::vector<Row> tableFromDatasheet(const std::vector<std::uint8_t>& bytes);

// Liest eine Datei auch dann, wenn Excel sie geöffnet hält.
std::vector<std::uint8_t> readFileShared(const std::filesystem::path& file);

}  // namespace rethink
