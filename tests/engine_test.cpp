// Unit-Tests des Layout-Kerns und des Datenblatts (ohne PowerPoint, ohne Dateien).
#include "cellinput.h"
#include "datasheet.h"
#include "engine/chart.h"

#include <objbase.h>

#include <cmath>
#include <cstdio>
#include <functional>
#include <stdexcept>
#include <string>

using namespace rethink;
using namespace rethink::engine;

namespace {

int failures = 0, checks = 0;

void check(bool ok, const std::string& what) {
    ++checks;
    if (!ok) {
        ++failures;
        std::printf("FEHL  %s\n", what.c_str());
    }
}

bool approx(double a, double b) { return std::fabs(a - b) < 1e-9; }

Cell num(double v) { return Cell{CellKind::Number, L"", v, std::nullopt}; }
Cell e() { return Cell{CellKind::String, L"e", 0, std::nullopt}; }
Cell none() { return Cell{}; }
Cell str(const wchar_t* s) { return Cell{CellKind::String, s, 0, std::nullopt}; }

// Kopfzeile mit n Kategorien + eine Datenzeile
DataItem single(std::vector<Cell> values) {
    DataItem item;
    item.name = L"Test";
    Row head{none()};
    for (size_t i = 0; i < values.size(); ++i) head.push_back(str((L"K" + std::to_wstring(i + 1)).c_str()));
    Row row{str(L"Reihe")};
    row.insert(row.end(), values.begin(), values.end());
    item.table = {head, row};
    return item;
}

double computedValue(const WaterfallColumn& c) {
    for (auto& s : c.segments)
        if (s.computed) return s.value;
    return NAN;
}

void testEAtEnd() {
    auto cols = computeWaterfall(single({num(10), num(5), num(-3), e()}), {});
    check(approx(cols[3].segments[0].value, 12), "e am Ende = 12");
    check(cols[3].anchored && approx(cols[3].segments[0].lo, 0) && approx(cols[3].segments[0].hi, 12), "e am Ende steht auf der Basislinie");
    check(approx(cols[1].segments[0].lo, 10) && approx(cols[1].segments[0].hi, 15), "Schritt +5 schwebt von 10 bis 15");
    check(approx(cols[2].segments[0].lo, 12) && approx(cols[2].segments[0].hi, 15), "Schritt -3 hängt von 15 bis 12");
    check(approx(cols[0].levelAfter, 10) && approx(cols[1].levelAfter, 15) && approx(cols[2].levelAfter, 12), "Verbinderhöhen 10/15/12");
}

void testEMid() {
    auto cols = computeWaterfall(single({num(10), num(5), e(), num(-3), e()}), {});
    check(approx(computedValue(cols[2]), 15), "Zwischensumme in der Mitte = 15");
    check(approx(computedValue(cols[4]), 12), "Endsumme nach Zwischensumme = 12");
    check(approx(cols[3].segments[0].lo, 12) && approx(cols[3].segments[0].hi, 15), "Schritt nach Zwischensumme setzt beim Stand an");
}

void testEFirst() {
    auto cols = computeWaterfall(single({e(), num(-2), num(-3), num(-4)}), {});
    check(approx(computedValue(cols[0]), 9), "e in erster Spalte rückwärts = 9");
    check(approx(cols[0].levelAfter, 9) && approx(cols[1].levelAfter, 7) && approx(cols[2].levelAfter, 4) && approx(cols[3].levelAfter, 0),
          "absteigend 9 → 7 → 4 → 0");
}

void testCrossingZero() {
    auto cols = computeWaterfall(single({num(5), num(-8), e()}), {});
    check(approx(cols[1].segments[0].lo, -3) && approx(cols[1].segments[0].hi, 5), "Schritt -8 kreuzt die Null (5 → -3)");
    check(approx(computedValue(cols[2]), -3) && approx(cols[2].segments[0].lo, -3) && approx(cols[2].segments[0].hi, 0),
          "negative Endsumme hängt unter der Basislinie");
}

void testTwoSegments() {
    DataItem item = single({num(10), num(5), e()});
    item.table.push_back(Row{str(L"Zweite"), none(), num(-2), none()});
    auto cols = computeWaterfall(item, {});
    check(cols[1].segments.size() == 2, "Säule 2 hat zwei Segmente");
    check(approx(cols[1].segments[0].lo, 10) && approx(cols[1].segments[0].hi, 15), "positives Segment 10–15");
    check(approx(cols[1].segments[1].lo, 8) && approx(cols[1].segments[1].hi, 10), "negatives Segment 8–10");
    check(approx(cols[1].levelAfter, 13) && approx(computedValue(cols[2]), 13), "Summe aus gemischten Vorzeichen = 13");
}

void testTwoEInColumn() {
    DataItem item = single({num(10), e()});
    item.table.push_back(Row{str(L"Zweite"), none(), e()});
    bool threw = false;
    try {
        computeWaterfall(item, {});
    } catch (const std::invalid_argument&) {
        threw = true;
    }
    check(threw, "zwei e in einer Säule werden abgelehnt");
}

void testTableOnlyRows() {
    DataItem item = single({num(10), num(5), e()});
    item.table.push_back(Row{str(L"Marge"), num(1), num(2), num(3)});
    ChartOptions o;
    o.tableRows = {2};
    o.dataTable = true;
    auto cols = computeWaterfall(item, o);
    check(cols[1].segments.size() == 1 && approx(computedValue(cols[2]), 15), "Tabellenzeile beeinflusst den Wasserfall nicht");

    Drawing d = layoutChart(item, o, Box{0, 0, 6000000, 4000000});
    int connectors = 0, tableCells = 0;
    bool computedShown = false;
    for (auto& p : d.prims) {
        if (p.role.starts_with("connector")) ++connectors;
        if (p.role.starts_with("table cell")) ++tableCells;
        if (p.role == "table cell r1 c3" && p.text == L"15") computedShown = true;
    }
    check(connectors == 2, "zwei Verbinder bei drei Säulen");
    check(tableCells == 6, "Datentabelle: 2 Zeilen × 3 Spalten");
    check(computedShown, "Datentabelle zeigt den berechneten e-Wert 15");
    for (auto& p : d.prims)
        check(p.x >= 0 && p.y >= 0 && p.x + std::max<Emu>(p.w, 0) <= 6000000 && p.y + std::max<Emu>(p.h, 0) <= 4000000 + 1,
              "Form liegt im Platzhalter: " + p.role);
}

void testFormat() {
    check(formatValue(1234.5, 1, L",", false) == L"1234,5", "Format 1234,5");
    check(formatValue(-0.04, 1, L",", false) == L"0,0" || formatValue(-0.04, 1, L",", false) == L"-0,0", "Format nahe -0");
    check(formatValue(-0.4, 0, L",", false) == L"0", "kein -0");
    check(formatValue(48, 0, L",", true) == L"48%", "Prozent");
}

void testDatasheetRoundtrip() {
    Cell red = num(-6);
    red.fill = 0xC00000;
    Cell pct{CellKind::Percentage, L"", 8.1, std::nullopt};
    Cell date{CellKind::Date, L"2024-03-01", 0, std::nullopt};
    std::vector<Row> table = {
        {none(), str(L"Start"), date, str(L"Ende")},
        {str(L"Wert"), num(120.5), red, e()},
        {},
        {str(L"Marge"), pct, none(), str(L"neu")},
    };
    const auto back = tableFromDatasheet(datasheetFromTable(table));
    check(back.size() == 4, "Datenblatt: 4 Zeilen inkl. Leerzeile");
    if (back.size() != 4) return;
    check(back[0].size() == 4 && back[0][0].kind == CellKind::Empty && back[0][1].text == L"Start", "Kopfzeile");
    check(back[0][2].kind == CellKind::Date && back[0][2].text == L"2024-03-01", "Datum bleibt Datum");
    check(back[1][1].kind == CellKind::Number && approx(back[1][1].number, 120.5), "Zahl");
    check(back[1][2].fill && *back[1][2].fill == 0xC00000 && approx(back[1][2].number, -6), "Zellfarbe bleibt erhalten");
    check(isComputedCell(back[1][3]), "e bleibt e");
    check(back[2].empty(), "Leerzeile bleibt leer");
    check(back[3][1].kind == CellKind::Percentage && approx(back[3][1].number, 8.1), "Prozent in Prozentpunkten");
    check(back[3].size() == 4 && back[3][2].kind == CellKind::Empty, "Lücke in der Zeile bleibt");
}

void testCellInput() {
    const std::wstring de = L",";
    auto kindIs = [&](const wchar_t* in, CellKind k) { return parseCellInput(in, de).kind == k; };
    check(kindIs(L"", CellKind::Empty) && kindIs(L"  ", CellKind::Empty), "Eingabe: leer");
    check(isComputedCell(parseCellInput(L"E", de)), "Eingabe: e");
    check(approx(parseCellInput(L"12,5", de).number, 12.5) && kindIs(L"12,5", CellKind::Number), "Eingabe: 12,5");
    check(approx(parseCellInput(L"12.5", de).number, 12.5), "Eingabe: 12.5 (Punkt)");
    check(approx(parseCellInput(L"1.234", de).number, 1234), "Eingabe: 1.234 als Tausender");
    check(approx(parseCellInput(L"1.234,5", de).number, 1234.5), "Eingabe: 1.234,5");
    check(approx(parseCellInput(L"-3", de).number, -3), "Eingabe: -3");
    check(kindIs(L"8,1 %", CellKind::Percentage) && approx(parseCellInput(L"8,1 %", de).number, 8.1), "Eingabe: 8,1 %");
    check(parseCellInput(L"01.03.2024", de).text == L"2024-03-01" && kindIs(L"2024-03-01", CellKind::Date), "Eingabe: Datum");
    check(kindIs(L"Umsatz", CellKind::String) && kindIs(L"12a", CellKind::String), "Eingabe: Text");
    check(displayCell(parseCellInput(L"12,5", de), de) == L"12,5", "Anzeige: 12,5");
    check(displayCell(parseCellInput(L"48%", de), de) == L"48%", "Anzeige: 48%");
    check(displayCell(parseCellInput(L"2024-03-01", de), de) == L"01.03.2024", "Anzeige: Datum");
}

}  // namespace

int main() {
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    testCellInput();
    testDatasheetRoundtrip();
    testEAtEnd();
    testEMid();
    testEFirst();
    testCrossingZero();
    testTwoSegments();
    testTwoEInColumn();
    testTableOnlyRows();
    testFormat();
    std::printf("%d von %d Prüfungen bestanden\n", checks - failures, checks);
    return failures == 0 ? 0 : 1;
}
