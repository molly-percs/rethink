#pragma once

#include <cstdint>
#include <optional>
#include <string>
#include <vector>

namespace rethink {

struct SheetCell {
    enum class Kind { Empty, Text, Number, Percent, Date } kind = Kind::Empty;
    std::wstring text;
    double value = 0;                   // Percent als Anteil (0.48), Date als Excel-Seriennummer
    std::optional<std::uint32_t> fill;  // Zellfarbe 0xRRGGBB
};

using SheetRow = std::vector<SheetCell>;

// Spaltenbuchstaben: 0 → A, 25 → Z, 26 → AA.
std::wstring columnName(size_t index);

// Minimale XLSX-Arbeitsmappe mit einem Blatt; Zeile 0 der Daten landet in Excel-Zeile 1.
std::vector<std::uint8_t> makeWorkbook(const std::vector<SheetRow>& rows, const std::wstring& sheetName = L"Sheet1");

// Liest das erste Blatt einer (z. B. von Excel gespeicherten) Arbeitsmappe:
// Zahlen, Text, Prozent- und Datumsformate sowie einfarbige Zellfüllungen.
// Leere Zeilen bleiben als leere Zeilen erhalten; Leerzellen am Zeilenende entfallen.
std::vector<SheetRow> readWorkbook(const std::vector<std::uint8_t>& bytes);

// Excel-Seriennummer (1900er-System) ↔ "YYYY-MM-DD".
double excelSerialFromIso(const std::wstring& iso);
std::wstring isoFromExcelSerial(double serial);

}  // namespace rethink
