// Unit-Tests des Layout-Kerns und des Datenblatts (ohne PowerPoint, ohne Dateien).
#include "cellinput.h"
#include "datasheet.h"
#include "engine/chart.h"
#include "engine/dates.h"
#include "engine/gantt_edit.h"

#include <objbase.h>

#include <algorithm>
#include <cmath>
#include <cstdio>
#include <functional>
#include <stdexcept>
#include <string>
#include <tuple>

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
    check(parseCellInput(L"1.3.26", de).text == L"2026-03-01" && kindIs(L"1.3.26", CellKind::Date), "Eingabe: Datum mit zweistelligem Jahr");
    check(kindIs(L"Umsatz", CellKind::String) && kindIs(L"12a", CellKind::String), "Eingabe: Text");
    check(displayCell(parseCellInput(L"12,5", de), de) == L"12,5", "Anzeige: 12,5");
    check(displayCell(parseCellInput(L"48%", de), de) == L"48%", "Anzeige: 48%");
    check(displayCell(parseCellInput(L"2024-03-01", de), de) == L"01.03.2024", "Anzeige: Datum");
}

void testStepAndDrag() {
    ChartOptions o;
    check(stepDecimals(num(45), o) == 0, "Schritt: ganze Zahl → 1er-Schritte");
    check(stepDecimals(num(-12.5), o) == 1, "Schritt: eine Nachkommastelle → 0,1");
    check(stepDecimals(num(3.25), o) == 2, "Schritt: zwei Nachkommastellen → 0,01");
    o.decimals = 0;
    check(stepDecimals(num(3.25), o) == 0, "Schritt: feste decimals-Option hat Vorrang");
    check(approx(dragValue(40, 25, 5, 0), 45), "Ziehen: 25 px bei 5 px/Einheit → +5");
    check(approx(dragValue(40, 27, 5, 0), 45) && approx(dragValue(40, 28, 5, 0), 46), "Ziehen: rundet auf ganze Schritte");
    check(approx(dragValue(12.5, -7, 10, 1), 11.8), "Ziehen: Nachkommastelle bleibt (12,5 → 11,8)");
    check(approx(dragValue(-3, 50, 10, 0), 2), "Ziehen: über die Null (−3 → 2)");
    check(dragValue(0.2, -2, 10, 0) == 0.0 && !std::signbit(dragValue(0.2, -2, 10, 0)), "Ziehen: kein −0");

    // Maßstab: alle Segmente teilen ihn; abgelesen am größten Betrag (die Endsumme 12)
    const Drawing d = layoutChart(single({num(10), num(5), num(-3), e()}), o, Box{0, 0, 6000000, 4000000});
    const double scale = emuPerUnit(d);
    bool consistent = scale > 0;
    for (auto& p : d.prims)
        if (p.kind == Prim::Kind::Rect && !p.flags.empty()) consistent = consistent && std::fabs(p.h - std::fabs(p.value) * scale) <= 1.0;
    check(consistent, "Maßstab: Segmenthöhe = |Wert| × EMU je Einheit");
    check(emuPerUnit(layoutChart(single({num(0), num(0)}), o, Box{0, 0, 6000000, 4000000})) == 0, "Maßstab: nur Nullen → 0");
}

bool inside(const Drawing& d, const Box& b) {
    for (auto& p : d.prims) {
        const Emu x0 = std::min(p.x, p.x + p.w), x1 = std::max(p.x, p.x + p.w), y0 = std::min(p.y, p.y + p.h), y1 = std::max(p.y, p.y + p.h);
        if (x0 < b.x - 1 || y0 < b.y - 1 || x1 > b.x + b.w + 1 || y1 > b.y + b.h + 1) {
            std::printf("      außerhalb: %s\n", p.role.c_str());
            return false;
        }
    }
    return true;
}

const Prim* find(const Drawing& d, const std::string& role) {
    for (auto& p : d.prims)
        if (p.role == role) return &p;
    return nullptr;
}

void testOrientation() {
    DataItem item = single({num(10), num(5), num(-3), e()});
    const Box box{0, 0, 6000000, 4000000};
    for (Orientation o : {Orientation::Up, Orientation::Down, Orientation::Right, Orientation::Left}) {
        ChartOptions opt;
        opt.orientation = o;
        opt.dataTable = true;
        const Drawing d = layoutChart(item, opt, box);
        const char* name = o == Orientation::Up ? "oben" : o == Orientation::Down ? "unten" : o == Orientation::Right ? "rechts" : "links";
        check(inside(d, box), std::string("Ausrichtung ") + name + ": alle Formen im Platzhalter");
        check(d.valueDirection == o, std::string("Ausrichtung ") + name + ": Richtung wachsender Werte vermerkt");
        const double scale = emuPerUnit(d);
        const bool horizontal = o == Orientation::Right || o == Orientation::Left;
        bool consistent = scale > 0;
        for (auto& p : d.prims)
            if (p.kind == Prim::Kind::Rect && !p.flags.empty())
                consistent = consistent && std::fabs((horizontal ? p.w : p.h) - std::fabs(p.value) * scale) <= 1.0;
        check(consistent, std::string("Ausrichtung ") + name + ": Segmentlänge = |Wert| × Maßstab");
        const Prim* first = find(d, "segment r1 c1");
        const Prim* second = find(d, "segment r1 c2");
        const Prim* base = find(d, "baseline");
        if (!first || !second || !base) {
            check(false, std::string("Ausrichtung ") + name + ": Segmente und Basislinie vorhanden");
            continue;
        }
        switch (o) {
            case Orientation::Up: check(first->y + first->h == base->y && second->x > first->x, "oben: erste Säule steht auf der Basislinie"); break;
            case Orientation::Down: check(first->y == base->y && second->y > first->y, "unten: Säulen hängen von der Basislinie nach unten"); break;
            case Orientation::Right: check(first->x == base->x && second->x > first->x && second->y > first->y, "rechts: Balken wachsen nach rechts, Kategorien nach unten"); break;
            case Orientation::Left:
                check(first->x + first->w == base->x && second->x < first->x && second->y > first->y, "links: Balken wachsen nach links");
                break;
        }
    }
}

void testDates() {
    check(daysFromCivil(1970, 1, 1) == 0 && daysFromCivil(2000, 3, 1) == 11017, "Tage seit 1970");
    const CivilDate leap = civilFromDays(daysFromCivil(2024, 2, 29));
    check(leap.year == 2024 && leap.month == 2 && leap.day == 29, "Schalttag hin und zurück");
    check(isoWeekday(daysFromCivil(2026, 10, 18)) == 7 && isoWeekday(daysFromCivil(2026, 10, 19)) == 1, "Wochentag (18.10.2026 = Sonntag)");
    check(isoWeek(daysFromCivil(2026, 1, 1)) == 1 && isoWeek(daysFromCivil(2021, 1, 3)) == 53 && isoWeek(daysFromCivil(2024, 12, 30)) == 1,
          "ISO-Kalenderwoche an Jahresgrenzen");
    check(parseIsoDate(L"2026-11-02") == static_cast<double>(daysFromCivil(2026, 11, 2)), "ISO-Datum");
    check(parseIsoDate(L"20261102") == parseIsoDate(L"2026-11-02") && parseIsoDate(L"2026") == static_cast<double>(daysFromCivil(2026, 1, 1)),
          "ISO-Datum kompakt und nur Jahr");
    check(approx(*parseIsoDate(L"2026-11-02T12:00") - *parseIsoDate(L"2026-11-02"), 0.5) && approx(*parseIsoDate(L"20261102T18Z") - *parseIsoDate(L"2026-11-02"), 0.75),
          "ISO-Datum mit Uhrzeit");
    check(!parseIsoDate(L"2026-02-30") && !parseIsoDate(L"2026-13") && !parseIsoDate(L"26-01-01") && !parseIsoDate(L""), "ungültige Daten abgelehnt");
    check(isoDate(*parseIsoDate(L"2026-11-02")) == L"2026-11-02" && isoDate(*parseIsoDate(L"2026-11-02T23:59")) == L"2026-11-02T23:59", "ISO-Ausgabe");

    const double jul5 = static_cast<double>(daysFromCivil(2024, 7, 5));
    check(formatDate(jul5, L"MMM 'yy", L"en") == L"Jul '24", "Format MMM 'yy");
    check(formatDate(static_cast<double>(daysFromCivil(2024, 10, 1)), L"\"Quarter\" QQQ", L"en") == L"Quarter IV", "Format \"Quarter\" QQQ");
    check(formatDate(jul5, L"ddd¶dd", L"en") == L"Fri\n05", "Format ddd¶dd (Zeilenumbruch)");
    check(formatDate(jul5, L"www", L"de") == L"KW 27" && formatDate(jul5, L"wwww", L"en") == L"Week 27", "Format Kalenderwoche");
    check(formatDate(jul5, L"dddd, d. MMMM yyyy", L"de") == L"Freitag, 5. Juli 2024", "Format deutsch ausgeschrieben");
    check(formatDate(static_cast<double>(daysFromCivil(2024, 3, 1)), L"MMM QQ \\Q", L"de") == L"Mär i Q", "Format römisch klein und \\-Zeichen");

    const double wed = static_cast<double>(daysFromCivil(2026, 11, 4));
    check(unitStart(wed, TimeUnit::Week, 1) == daysFromCivil(2026, 11, 2), "Wochenanfang Montag");
    check(unitStart(wed, TimeUnit::Week, 7) == daysFromCivil(2026, 11, 1), "Wochenanfang Sonntag");
    check(unitStart(wed, TimeUnit::Quarter, 1) == daysFromCivil(2026, 10, 1) && nextUnit(daysFromCivil(2026, 10, 1), TimeUnit::Quarter) == daysFromCivil(2027, 1, 1),
          "Quartal");
    check(addUnits(static_cast<double>(daysFromCivil(2026, 1, 31)), TimeUnit::Month, 1) == static_cast<double>(daysFromCivil(2026, 2, 28)),
          "Monat weiter: Tag wird begrenzt");
    check(snapToUnit(wed + 0.4, TimeUnit::Day, 1) == wed && snapToUnit(wed + 0.6, TimeUnit::Day, 1) == wed + 1, "Raster Tag");
    check(snapToUnit(wed, TimeUnit::Week, 1) == static_cast<double>(daysFromCivil(2026, 11, 2)), "Raster Woche");
}

Cell date(const wchar_t* iso) { return Cell{CellKind::Date, iso, 0, std::nullopt}; }

// Zeitplan wie im Datenblatt: Überschriften, Meilensteinzeile, drei Aktivitäten
DataItem plan() {
    DataItem item;
    item.name = L"Plan";
    item.table = {
        {str(L"Aktivität"), str(L"Start"), str(L"Ende"), str(L"Meilenstein"), str(L"Verantwortlich"), str(L"Bemerkung")},
        {str(L"Meilensteine"), none(), none(), date(L"2027-01-15")},
        {str(L"Analyse"), date(L"2026-11-02"), date(L"2026-11-20"), none()},
        {str(L"Konzept"), date(L"2026-11-16"), date(L"2026-12-11"), date(L"2026-12-11")},
        {str(L"Umsetzung"), date(L"2026-12-07"), date(L"2027-01-29"), none(), str(L"Team B")},
    };
    return item;
}

ChartOptions ganttOptions() {
    ChartOptions o;
    o.type = ChartType::Gantt;
    o.gantt.start = L"2026-11-01";
    o.gantt.end = L"2027-01-31";
    o.gantt.dateFormat = L"dd.MM.";  // kurze Beschriftungen in den Prüfungen
    return o;
}

void testGanttLayout() {
    const Box box{0, 0, 9000000, 4000000};
    DataItem item = plan();
    ChartOptions o = ganttOptions();
    const Drawing d = layoutChart(item, o, box);
    check(inside(d, box), "Gantt: alle Formen im Platzhalter");
    check(d.axis.valid && d.axis.t0 == static_cast<double>(daysFromCivil(2026, 11, 1)) && d.axis.t1 == static_cast<double>(daysFromCivil(2027, 2, 1)),
          "Gantt: Bereich 01.11.2026 bis einschließlich 31.01.2027");
    const auto scales = ganttScales(d.axis.t0, d.axis.t1, d.axis.w, o);
    check(scales.size() == 2 && scales[0] == TimeUnit::Month && scales[1] == TimeUnit::Week, "Gantt: drei Monate auf 25 cm → Monats- und Wochenskala");
    const Prim* month = find(d, "scale month 1");
    check(month && month->text == L"Nov" && month->bold && month->align == Align::Left, "Gantt: Monat kurz, fett, linksbündig wie im Original („Nov“)");
    // Wochen sind hier zu schmal für "dd.MM.": Kalenderwochen ab Montag 02.11. (KW 45)
    check(find(d, "scale week 2") && find(d, "scale week 2")->text == L"45", "Gantt: Wochenskala zeigt KW 45 für Montag 02.11.");
    const Prim* bar = find(d, "bar r2");
    check(bar && bar->flags == "bar" && std::llabs(bar->x - std::llround(d.axis.xOf(*parseIsoDate(L"2026-11-02")))) <= 1 &&
              std::llabs(bar->x + bar->w - std::llround(d.axis.xOf(*parseIsoDate(L"2026-11-20")))) <= 1,
          "Gantt: Balken von Start bis Ende (Ende 00:00 Uhr)");
    check(find(d, "activity r2") && find(d, "activity r2")->text == L"Analyse", "Gantt: Aktivitätsbeschriftung");
    check(find(d, "headline activity") && find(d, "headline activity")->bold && find(d, "headline responsible"),
          "Gantt: Spaltenüberschriften fett; Spalte „Verantwortlich“, weil es einen Eintrag gibt");
    check(!find(d, "headline remark"), "Gantt: leere Spalte „Bemerkung“ ist ausgeblendet");
    check(find(d, "responsible r4") && find(d, "responsible r4")->text == L"Team B", "Gantt: Verantwortlich-Eintrag");
    const Prim* ms = find(d, "milestone r3");
    check(ms && ms->kind == Prim::Kind::Shape && ms->shape == "triangle" && ms->rotation == 0 &&
              std::llabs(ms->x + ms->w / 2 - std::llround(d.axis.xOf(*parseIsoDate(L"2026-12-11")))) <= 1,
          "Gantt: Meilenstein ▲ am Tagesanfang (ohne Tagesskala)");
    check(find(d, "milestone label r3") && find(d, "milestone label r3")->text == L"11.12.", "Gantt: Meilenstein mit Datum");
    const Prim* key = find(d, "milestone r1");
    const Prim* lastBar = find(d, "bar r4");
    check(key && lastBar && key->y > lastBar->y + lastBar->h, "Gantt: Meilensteinzeile unter dem Diagramm");
    check(find(d, "activity r1") && find(d, "activity r1")->text == L"Meilensteine", "Gantt: Beschriftung der Meilensteinzeile");
    int separators = 0, dashed = 0;
    for (auto& p : d.prims)
        if (p.role.starts_with("separator")) ++separators, dashed += p.dotted;
    check(separators > 0 && dashed > 0 && dashed < separators, "Gantt: Trennlinien – Monate durchgezogen, Wochen gepunktet");
    bool weekends = false, plain = true;
    for (auto& p : d.prims) plain = plain && !p.role.starts_with("weekend");
    check(plain, "Gantt: Wochenenden wie im Original zunächst ohne Grau");
    ChartOptions grey = o;
    grey.gantt.weekends = WeekendMode::Shaded;
    for (auto& p : layoutChart(item, grey, box).prims) weekends = weekends || p.role.starts_with("weekend");
    check(weekends, "Gantt: Wochenenden auch in der Wochenansicht leicht grau");
    // Aussehen wie im Original: Rahmen der Zeitachse, keine Teilstriche, ▲ der Meilensteinzeile mit Datum darunter und Linie
    check(find(d, "frame left") && find(d, "frame right") && !find(d, "scale week tick 2"), "Standard: Rahmen links/rechts, Skalen ohne Teilstriche");
    const Prim* keyMark = find(d, "milestone r1");
    const Prim* keyDate = find(d, "milestone label r1");
    const Prim* keyLine = find(d, "milestone line r1");
    check(keyMark && keyDate && keyDate->y >= keyMark->y + keyMark->h && std::llabs(keyDate->x + keyDate->w / 2 - (keyMark->x + keyMark->w / 2)) <= 1,
          "Standard: Datum mittig unter dem ▲ der Meilensteinzeile");
    check(keyLine && keyLine->dashed && keyLine->y == find(d, "frame left")->y, "Standard: gestrichelte Meilensteinlinie bis zur Oberkante");
    check(find(d, "bar r2") && find(d, "bar r2")->fill.scheme == L"tx1", "Standard: Balken in Textfarbe (schwarz) wie im Original");
    {
        ChartOptions lang = ganttOptions();
        lang.gantt.dateFormat.clear();
        check(find(layoutChart(item, lang, box), "milestone label r1")->text == L"15.01.2027", "Datumsformat nach Sprache: dd.MM.yyyy");
        lang.language = L"en";
        check(find(layoutChart(item, lang, box), "milestone label r1")->text == L"1/15/2027", "Datumsformat nach Sprache: M/d/yyyy");
    }
    {
        // Aussehen je Element: gestrichelter Balken (wie „Dashed Bar“), Füllfarbe eines Meilensteins
        ChartOptions styled = ganttOptions();
        styled.gantt.rowStyles = {{2, {L"bg1", L"tx1", 1.5, true}, {}}, {3, {}, {L"accent2", L"", 0, false}}};
        const Drawing s = layoutChart(item, styled, box);
        const Prim* dashedBar = find(s, "bar r2");
        check(dashedBar && dashedBar->fill.scheme == L"bg1" && dashedBar->stroke.scheme == L"tx1" && dashedBar->dashed && dashedBar->strokePt == 1.5,
              "Stil je Zeile: gestrichelter Balken");
        check(find(s, "milestone r3") && find(s, "milestone r3")->fill.scheme == L"accent2", "Stil je Zeile: Farbe des Meilensteins");
        check(find(s, "bar r3") && find(s, "bar r3")->stroke.kind == Color::Kind::None, "ohne Stil: kein Umriss");
    }
    check(d.prims.front().role == "background" && d.prims.front().transparency == 1, "Gantt: unsichtbare Fläche hinter allem (Klick trifft das Diagramm)");
    {
        ChartOptions years = ganttOptions();
        years.gantt.end = L"2028-12-31";
        bool none = true;
        for (auto& p : layoutChart(item, years, box).prims) none = none && !p.role.starts_with("weekend");
        check(none, "Gantt: über zwei Jahre keine Wochenenden (zu schmal)");
    }

    // Zwei Wochen: Tagesskala, Wochenenden schattiert, Meilensteine in der Tagesmitte
    o.gantt.start = L"2026-12-01";
    o.gantt.end = L"2026-12-14";
    o.gantt.weekends = WeekendMode::Shaded;
    const Drawing days = layoutChart(item, o, box);
    check(days.axis.dayScale && find(days, "scale day 1"), "Gantt: zwei Wochen → Tagesskala");
    bool weekend = false;
    for (auto& p : days.prims) weekend = weekend || p.role.starts_with("weekend");
    check(weekend, "Gantt: Wochenenden schattiert");
    const Prim* mid = find(days, "milestone r3");
    check(mid && std::llabs(mid->x + mid->w / 2 - std::llround(days.axis.xOf(*parseIsoDate(L"2026-12-11") + 0.5))) <= 1,
          "Gantt: Meilenstein in der Tagesmitte (Tagesskala sichtbar)");
    const Prim* clipped = find(days, "bar r3");
    check(clipped && clipped->x == days.axis.x0 && find(days, "bar r3 from") && !find(days, "bar r3 to"),
          "Gantt: Balken beginnt vor dem Bereich → Pfeil am Anfang, keiner am Ende");
    check(find(days, "bar r4") && find(days, "bar r4 to"), "Gantt: Balken über das Bereichsende hinaus bekommt einen Pfeil");
    check(find(days, "bar r2") == nullptr, "Gantt: Balken außerhalb des Bereichs ist ausgeblendet");

    // Ausgeblendete Wochenenden: Samstag und Sonntag haben keine Breite
    o.gantt.weekends = WeekendMode::Hidden;
    const GanttAxis hidden = ganttAxis(item, o, box);
    const double sat = *parseIsoDate(L"2026-12-05"), mon = *parseIsoDate(L"2026-12-07");
    check(std::fabs(hidden.xOf(sat) - hidden.xOf(mon)) < 1 && hidden.xOf(sat) > hidden.xOf(sat - 1), "Gantt: ausgeblendete Wochenenden ohne Breite");
    check(std::fabs(hidden.dateAt(hidden.xOf(*parseIsoDate(L"2026-12-03T06:00"))) - *parseIsoDate(L"2026-12-03T06:00")) < 1e-3,
          "Gantt: Achse umkehrbar (dateAt ∘ xOf)");

    // Ohne Bereich: an die Daten angepasst (ganze Monate); ohne Daten: Referenzmonat + zwei
    ChartOptions fit;
    fit.type = ChartType::Gantt;
    const auto range = ganttRange(item, fit);
    check(range.first == static_cast<double>(daysFromCivil(2026, 11, 1)) && range.second == static_cast<double>(daysFromCivil(2027, 2, 1)),
          "Gantt: Vorgabebereich = Daten auf Monate erweitert");
    DataItem empty;
    empty.table = {{str(L"Aktivität")}};
    fit.gantt.referenceDay = static_cast<double>(daysFromCivil(2026, 10, 18));
    const auto fresh = ganttRange(empty, fit);
    check(fresh.first == static_cast<double>(daysFromCivil(2026, 10, 1)) && fresh.second == static_cast<double>(daysFromCivil(2027, 1, 1)),
          "Gantt: neues Diagramm zeigt den aktuellen und die zwei folgenden Monate");

    // Frei gesetzte Elemente
    ChartOptions items = ganttOptions();
    items.gantt.items = {
        {GanttItem::Kind::Chevron, 2, L"2026-11-23", L"2026-12-04", L"Phase 2", std::nullopt},
        {GanttItem::Kind::Shade, 0, L"2026-12-24", L"2027-01-02", L"", std::nullopt},
        {GanttItem::Kind::Bracket, 4, L"2027-01-04", L"2027-01-16", L"Test", std::nullopt},
        {GanttItem::Kind::Milestone, 1, L"2026-12-18", L"", L"", 0xC00000},
    };
    const Drawing withItems = layoutChart(item, items, box);
    const Prim* chevron = find(withItems, "item 0");
    check(chevron && chevron->shape == "chevron" && chevron->text == L"Phase 2", "Gantt: Chevron mit Text");
    const Prim* shade = find(withItems, "item 1");
    check(shade && shade->flags == "shade" && find(withItems, "bar r2") && shade < find(withItems, "bar r2") && shade < find(withItems, "separator month 1"),
          "Gantt: Schattierung liegt hinter Trennlinien und Balken");
    check(find(withItems, "item 1 label") && find(withItems, "item 1 label")->text == L"24.12.–01.01.", "Gantt: Schattierung mit Datumsbereich");
    check(find(withItems, "item 2") && find(withItems, "item 2 label")->text == L"Test", "Gantt: Klammer mit Text");
    const Prim* red = find(withItems, "item 3");
    check(red && red->fill.kind == Color::Kind::Rgb && red->fill.rgb == 0xC00000, "Gantt: Meilenstein mit eigener Farbe");
    check(inside(withItems, box), "Gantt mit Elementen: alle Formen im Platzhalter");

    // Datumsbeschriftung an Balken: letzter eingeschlossener Tag
    ChartOptions labels = ganttOptions();
    labels.gantt.barLabels = BarLabel::Range;
    const Drawing labelled = layoutChart(item, labels, box);
    check(find(labelled, "bar r2 label") && find(labelled, "bar r2 label")->text == L"02.11.–19.11.", "Gantt: Balkenbeschriftung bis zum letzten Tag");

    bool tooSmall = false;
    try {
        layoutChart(item, o, Box{0, 0, 9000000, 600000});
    } catch (const std::invalid_argument&) {
        tooSmall = true;
    }
    check(tooSmall, "Gantt: zu niedriger Platzhalter wird gemeldet");
}

// Kategorien, Formen aus dem Datenblatt, Verbindungen und Designs
void testGanttCategories() {
    const Box box{0, 0, 9000000, 4000000};
    DataItem item = plan();
    for (auto& row : item.table) row.resize(gantt::kColumns);
    item.table[0][gantt::kCategoryCol] = str(L"Kategorie");
    item.table[2][gantt::kCategoryCol] = str(L"Planung");
    item.table[3][gantt::kCategoryCol] = str(L"Planung");
    item.table[4][gantt::kCategoryCol] = str(L"Bau");
    item.table[3][gantt::kPredecessorCol] = str(L"Analyse");
    item.table[4][gantt::kPredecessorCol] = str(L"konzept");  // Groß-/Kleinschreibung egal
    item.table[3][gantt::kShapeCol] = str(L"Chevron");
    item.table[3][gantt::kMarkerCol] = str(L"Dreieck unten");
    check(ganttCategories(item) == std::vector<std::wstring>{L"Planung", L"Bau"}, "Kategorien in Reihenfolge des Auftretens");
    check(canonicalBarShape(L"Fünfeck") == L"pentagon" && canonicalMarker(L"Stern") == L"star" && canonicalBarShape(L"Kreis").empty(),
          "Formnamen deutsch/englisch");

    ChartOptions o = ganttOptions();
    Drawing d = layoutChart(item, o, box);
    check(inside(d, box), "Kategorien: alle Formen im Platzhalter");
    const Prim* head = find(d, "headline category");
    const Prim* activity = find(d, "headline activity");
    check(head && activity && head->x < activity->x, "Kategoriespalte steht links von den Aktivitäten");
    const Prim* planung = find(d, "category label r2");
    const Prim* bau = find(d, "category label r4");
    check(planung && planung->text == L"Planung" && planung->bold && bau && bau->text == L"Bau", "eine fette Beschriftung je Kategoriegruppe");
    check(find(d, "category r2") && find(d, "category line r4") && !find(d, "category line r2"), "Standard: Klammer je Gruppe, Linie zwischen den Gruppen");
    check(planung && find(d, "bar r3") && planung->h > find(d, "bar r3")->h * 2, "Gruppe überspannt ihre Zeilen");
    const Prim* chevron = find(d, "bar r3");
    check(chevron && chevron->kind == Prim::Kind::Shape && chevron->shape == "chevron", "Form aus Spalte 8: Chevron");
    const Prim* triangle = find(d, "milestone r3");
    check(triangle && triangle->shape == "triangle" && triangle->rotation == 180, "Marker aus Spalte 9: Dreieck unten (▼)");
    check(find(d, "link r3 1 arrow") && find(d, "link r4 1 arrow"), "Vorgänger als Pfeile verbunden");
    o.gantt.links = false;
    check(!find(layoutChart(item, o, box), "link r3 1 arrow"), "Verbindungen ausschaltbar");

    // Designs: Farben sind Designfarben, Modern mit Kopfband, abgerundeten Balken, Kategorie-Blöcken und Kategoriefarben
    ChartOptions modern = ganttOptions();
    modern.gantt.design = GanttDesign::Modern;
    const Drawing m = layoutChart(item, modern, box);
    check(find(m, "header band") && find(m, "category r2") && find(m, "category r2")->kind == Prim::Kind::Rect, "Modern: Kopfband und Kategorie-Blöcke");
    const Prim* rounded = find(m, "bar r2");
    check(rounded && rounded->shape == "roundRect" && rounded->fill.scheme == L"accent1", "Modern: abgerundete Balken in der Farbe ihrer Kategorie");
    check(find(m, "bar r4") && find(m, "bar r4")->fill.scheme == L"accent2", "zweite Kategorie: zweite Akzentfarbe");
    bool themed = true;
    for (auto& p : m.prims)
        for (const Color* c : {&p.fill, &p.stroke, &p.textColor}) themed = themed && c->kind != Color::Kind::Rgb;
    check(themed, "Modern: nur Designfarben (folgen dem Folienmaster)");
    ChartOptions thin = modern;
    thin.gantt.barHeight = 0.3;
    check(find(layoutChart(item, thin, box), "bar r2")->h < rounded->h, "Balkenhöhe einstellbar");
    thin.gantt.categoryStyle = CategoryStyle::Text;
    check(!find(layoutChart(item, thin, box), "category r2") && find(layoutChart(item, thin, box), "category label r2"), "Kategoriestil nur Text");
    ChartOptions contrast = ganttOptions();
    contrast.gantt.design = GanttDesign::Contrast;
    const Drawing c = layoutChart(item, contrast, box);
    check(find(c, "header band") && find(c, "header band")->fill.scheme == L"tx2" && find(c, "headline activity")->textColor.scheme == L"bg1",
          "Kontrast: dunkles Kopfband mit heller Schrift");
    check(find(c, "row line r3") && !find(c, "row line r4"), "Kontrast: Zeilenlinien, an Kategoriegrenzen die Kategorielinie");
    ChartOptions minimal = ganttOptions();
    minimal.gantt.design = GanttDesign::Minimal;
    bool noSeparators = true;
    for (auto& p : layoutChart(item, minimal, box).prims) noSeparators = noSeparators && !p.role.starts_with("separator");
    check(noSeparators, "Minimal: keine senkrechten Trennlinien");
    minimal.gantt.separators = true;
    check(find(layoutChart(item, minimal, box), "separator month 1"), "Trennlinien einschaltbar trotz Design");

    // Nach Kategorie sortieren: gewählte Reihenfolge, stabil, Elemente wandern mit
    ChartOptions so = ganttOptions();
    so.gantt.items = {{GanttItem::Kind::Bar, 4, L"2026-12-14", L"2026-12-18", L"", std::nullopt}};
    so.gantt.shadedRows = {2};
    DataItem sorted = item;
    sortGanttByCategory(sorted, so, {L"Bau", L"Planung"});
    check(sorted.table[2][0].text == L"Umsetzung" && sorted.table[3][0].text == L"Analyse" && sorted.table[4][0].text == L"Konzept",
          "Sortieren: Bau zuerst, Planung stabil dahinter");
    check(so.gantt.items[0].row == 2 && so.gantt.shadedRows[0] == 3, "Sortieren: Elemente und Schattierungen folgen ihrer Zeile");
    // Zeile im Diagramm ziehen: innerhalb der Kategorie, dann in eine andere (übernimmt deren Kategorie)
    {
        DataItem drag = item;  // Analyse, Konzept (Planung), Umsetzung (Bau)
        ChartOptions dragOptions = ganttOptions();
        dragOptions.gantt.rowStyles = {{2, {L"accent3", L"", 0, false}, {}}};
        check(moveGanttRowTo(drag, dragOptions, 2, 3) && drag.table[2][0].text == L"Konzept" && drag.table[3][0].text == L"Analyse" &&
                  drag.table[3][gantt::kCategoryCol].text == L"Planung" && dragOptions.gantt.rowStyles[0].row == 3,
              "Zeile innerhalb der Kategorie nach unten gezogen, Stil folgt");
        check(moveGanttRowTo(drag, dragOptions, 3, 4) && drag.table[4][0].text == L"Analyse" && drag.table[4][gantt::kCategoryCol].text == L"Bau" &&
                  drag.table[3][0].text == L"Umsetzung",
              "Zeile in die Kategorie Bau gezogen: übernimmt die Kategorie im Datenblatt");
        check(!moveGanttRowTo(drag, dragOptions, 4, 4) && !moveGanttRowTo(drag, dragOptions, 1, 3), "Ablage an derselben Stelle bzw. Meilensteinzeile: nichts");
    }
    check(toggleGanttLink(sorted, 4, L"Umsetzung") && sorted.table[4][gantt::kPredecessorCol].text == L"Analyse; Umsetzung", "Verbindung hinzufügen");
    check(!toggleGanttLink(sorted, 4, L"Analyse") && sorted.table[4][gantt::kPredecessorCol].text == L"Umsetzung", "Verbindung entfernen");
}

void testGanttEdit() {
    const Box box{0, 0, 9000000, 4000000};
    DataItem item = plan();
    ChartOptions o = ganttOptions();
    const GanttAxis axis = ganttAxis(item, o, box);
    const double perDay = axis.emuPerDay();

    GanttTarget bar = ganttTarget(item, o, "bar r2");
    check(bar.kind == GanttTarget::Kind::Bar && bar.row == 2 && bar.start == *parseIsoDate(L"2026-11-02"), "Ziel: Balken aus dem Datenblatt");
    auto [s, e] = ganttDragDates(axis, bar, GanttGrip::End, perDay * 3.4, TimeUnit::Day);
    check(s == bar.start && e == *parseIsoDate(L"2026-11-23"), "Ende ziehen: +3 Tage (gerastet)");
    std::tie(s, e) = ganttDragDates(axis, bar, GanttGrip::Move, -perDay * 2.2, TimeUnit::Day);
    check(s == *parseIsoDate(L"2026-10-31") && e - s == bar.end - bar.start, "Verschieben: −2 Tage, Dauer bleibt");
    std::tie(s, e) = ganttDragDates(axis, bar, GanttGrip::Start, perDay * 40, TimeUnit::Day);
    check(s == bar.end - 1 && e == bar.end, "Start über das Ende hinaus: bleibt einen Tag davor");
    std::tie(s, e) = ganttDragDates(axis, bar, GanttGrip::Start, perDay * 4, TimeUnit::Week);
    check(s == *parseIsoDate(L"2026-11-09"), "Start ziehen mit Wochenraster: nächster Montag");
    setGanttDates(item, o, bar, *parseIsoDate(L"2026-11-03"), *parseIsoDate(L"2026-11-24"));
    check(item.table[2][1].kind == CellKind::Date && item.table[2][1].text == L"2026-11-03" && item.table[2][2].text == L"2026-11-24",
          "Neue Daten stehen im Datenblatt");

    const GanttTarget ms = ganttTarget(item, o, "milestone r3");
    std::tie(s, e) = ganttDragDates(axis, ms, GanttGrip::Move, perDay * 5.2, TimeUnit::Day);
    check(ms.kind == GanttTarget::Kind::Milestone && s == *parseIsoDate(L"2026-12-16"), "Meilenstein ziehen: +5 Tage");
    check(ganttTarget(item, o, "bar r9").kind == GanttTarget::Kind::None && ganttTarget(item, o, "bar r2 label").kind == GanttTarget::Kind::None,
          "Ziel: unbekannte Rollen");

    const auto [t0, t1] = ganttRangeDrag(axis, true, -perDay * 12, TimeUnit::Week);
    check(t0 == *parseIsoDate(L"2026-10-19") && t1 == axis.t1, "Skala links ziehen: Wochenraster");
    const auto [u0, u1] = ganttRangeDrag(axis, false, -perDay * 200, TimeUnit::Month);
    check(u0 == axis.t0 && u1 == *parseIsoDate(L"2026-12-01"), "Skala rechts zu weit: mindestens eine Einheit bleibt");
    check(ganttTooltip(*parseIsoDate(L"2026-11-04"), L"de") == L"Mi 04.11.2026 · KW 45", "Kurzinfo beim Ziehen");

    ChartOptions fit = o;
    check(fitGanttRange(item, fit) && fit.gantt.start == L"2026-11-03" && fit.gantt.end == L"2027-01-28", "Skala an Daten anpassen");

    // Zeilen und Elemente
    o.gantt.items = {{GanttItem::Kind::Bar, 3, L"2026-12-14", L"2026-12-18", L"", std::nullopt}};
    const int added = insertGanttRow(item, o, 3, L"Neu");
    check(added == 3 && item.table[3][0].text == L"Neu" && item.table[4][0].text == L"Konzept" && o.gantt.items[0].row == 4,
          "Zeile einfügen: Elemente darunter wandern mit");
    check(moveGanttRow(item, o, 4, -1) && item.table[3][0].text == L"Konzept" && o.gantt.items[0].row == 3, "Zeile nach oben: Element folgt");
    check(!moveGanttRow(item, o, 2, -1), "Zeile über die erste Aktivität hinaus: nicht möglich");
    duplicateGanttRow(item, o, 3, false);
    check(item.table[4][0].text == L"Konzept" && o.gantt.items.size() == 2 && o.gantt.items[1].row == 4, "Zeile duplizieren mit Elementen");
    deleteGanttRow(item, o, 4);
    check(item.table[4][0].text == L"Neu" && o.gantt.items.size() == 1, "Zeile löschen entfernt ihre Elemente");

    addGanttItem(item, o, GanttItem::Kind::Bar, 4, *parseIsoDate(L"2026-12-01T15:00"), 7);
    check(item.table[4].size() > 2 && item.table[4][1].text == L"2026-12-01" && item.table[4][2].text == L"2026-12-08", "Neuer Balken in leerer Zeile → Datenblatt");
    addGanttItem(item, o, GanttItem::Kind::Bar, 4, *parseIsoDate(L"2026-12-20"), 7);
    check(o.gantt.items.size() == 2 && o.gantt.items[1].row == 4 && o.gantt.items[1].start == L"2026-12-20", "Zweiter Balken in der Zeile → freies Element");
    deleteGanttTarget(item, o, ganttTarget(item, o, "item 1"));
    deleteGanttTarget(item, o, ganttTarget(item, o, "bar r4"));
    check(o.gantt.items.size() == 1 && item.table[4][1].kind == CellKind::Empty, "Elemente löschen");
}

}  // namespace

int main() {
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    testCellInput();
    testStepAndDrag();
    testOrientation();
    testDates();
    testGanttLayout();
    testGanttCategories();
    testGanttEdit();
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
