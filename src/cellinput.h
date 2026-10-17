#pragma once

#include "ppttc.h"

#include <string>

namespace rethink {

// Eingabe einer Datenblattzelle deuten (wie in einer Tabellenkalkulation):
//  ""            → leer
//  "e" / "E"     → berechnetes Segment
//  "12,5" "-3"   → Zahl (Dezimaltrennzeichen des Gebietsschemas; "." wird ebenfalls akzeptiert,
//                   ein Tausendertrennzeichen vor genau drei Ziffern wird ignoriert)
//  "48%" "8,1 %" → Prozent (in Prozentpunkten)
//  "01.03.2024" "2024-03-01" → Datum
//  sonst         → Text
Cell parseCellInput(const std::wstring& text, const std::wstring& decimalSeparator);

// Anzeige einer Zelle im Datenblatt (Zahlen mit Gebietsschema-Trennzeichen, Datum als TT.MM.JJJJ).
std::wstring displayCell(const Cell& cell, const std::wstring& decimalSeparator);

}  // namespace rethink
