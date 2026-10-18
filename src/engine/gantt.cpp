#include "chart.h"
#include "dates.h"
#include "draw.h"

#include <algorithm>
#include <cmath>
#include <format>
#include <stdexcept>

namespace rethink::engine {
namespace {

using namespace draw;
using namespace gantt;

constexpr std::uint32_t kGrid = 0xBFBFBF, kWeekend = 0xF2F2F2, kRowShade = 0xF2F2F2, kShade = 0xDDEBF7, kToday = 0xC00000;

const Cell& at(const DataItem& item, size_t r, size_t c) {
    static const Cell empty;
    return r < item.table.size() && c < item.table[r].size() ? item.table[r][c] : empty;
}

std::wstring textAt(const DataItem& item, size_t r, size_t c) {
    const Cell& cell = at(item, r, c);
    return cell.kind == CellKind::Date ? std::wstring() : cellText(cell);
}

// Zeitanteil von [a, b) ohne Wochenenden (Tage; b < a ergibt einen negativen Wert).
double workTime(double a, double b, int first, int last) {
    if (b < a) return -workTime(b, a, first, last);
    double sum = 0;
    for (long long day = static_cast<long long>(std::floor(a)); static_cast<double>(day) < b; ++day) {
        if (isWeekend(day, first, last)) continue;
        sum += std::min(b, static_cast<double>(day + 1)) - std::max(a, static_cast<double>(day));
    }
    return sum;
}

// Vorgabeformate einer Skala, lang → kurz
const std::vector<const wchar_t*>& defaultFormats(TimeUnit unit) {
    static const std::vector<const wchar_t*> day{L"ddd dd", L"dd", L"d"}, week{L"dd.MM.", L"ww"}, month{L"MMMM", L"MMM", L"MMMMM"},
        quarter{L"\"Q\"Q yyyy", L"\"Q\"Q", L"Q"}, year{L"yyyy", L"yy"};
    switch (unit) {
        case TimeUnit::Day: return day;
        case TimeUnit::Week: return week;
        case TimeUnit::Month: return month;
        case TimeUnit::Quarter: return quarter;
        case TimeUnit::Year: return year;
    }
    return day;
}

const std::wstring& customFormat(const GanttOptions& g, TimeUnit unit) {
    switch (unit) {
        case TimeUnit::Day: return g.dayFormat;
        case TimeUnit::Week: return g.weekFormat;
        case TimeUnit::Month: return g.monthFormat;
        case TimeUnit::Quarter: return g.quarterFormat;
        case TimeUnit::Year: return g.yearFormat;
    }
    return g.dayFormat;
}

// Breitestes Skalenformat, dessen Beschriftungen in die kürzeste Einheit passen; leer, wenn keins passt.
std::wstring scaleFormat(TimeUnit unit, double unitEmu, const ChartOptions& o) {
    const std::wstring& custom = customFormat(o.gantt, unit);
    const Emu pad = textHeight(o.fontPt) / 3;
    // Probedatum mit breiten Namen (Mittwoch, September, KW 38)
    const double probe = static_cast<double>(daysFromCivil(2025, 9, 17));
    if (!custom.empty()) return custom;
    for (const wchar_t* f : defaultFormats(unit))
        if (textWidth(formatDate(probe, f, o.language), o.fontPt) + pad <= unitEmu) return f;
    return {};
}

double minUnitDays(TimeUnit unit) {
    switch (unit) {
        case TimeUnit::Month: return 28;
        case TimeUnit::Quarter: return 90;
        case TimeUnit::Year: return 365;
        default: return unitDays(unit);
    }
}

TimeUnit parentUnit(TimeUnit unit) {
    switch (unit) {
        case TimeUnit::Day: return TimeUnit::Month;
        case TimeUnit::Week: return TimeUnit::Month;
        case TimeUnit::Month: return TimeUnit::Year;
        case TimeUnit::Quarter: return TimeUnit::Year;
        case TimeUnit::Year: return TimeUnit::Year;
    }
    return TimeUnit::Year;
}

std::string unitKey(TimeUnit unit) {
    switch (unit) {
        case TimeUnit::Day: return "day";
        case TimeUnit::Week: return "week";
        case TimeUnit::Month: return "month";
        case TimeUnit::Quarter: return "quarter";
        case TimeUnit::Year: return "year";
    }
    return "day";
}

std::wstring headline(const DataItem& item, int col, const wchar_t* de, const wchar_t* en, const std::wstring& language) {
    const std::wstring s = textAt(item, kHeaderRow, col);
    return s.empty() ? (language == L"en" ? en : de) : s;
}

// Letzter eingeschlossener Tag eines Endes (Ende um 00:00 Uhr schließt diesen Tag nicht ein).
double lastDay(double end) { return std::fabs(end - std::round(end)) < 1e-9 ? end - 1 : std::floor(end); }

std::wstring rangeText(double start, double end, const std::wstring& format, const std::wstring& language, BarLabel what) {
    const std::wstring a = formatDate(start, format, language), b = formatDate(lastDay(end), format, language);
    switch (what) {
        case BarLabel::Start: return a;
        case BarLabel::End: return b;
        default: return a == b ? a : a + L"–" + b;
    }
}

struct Layout {
    const DataItem& item;
    const ChartOptions& o;
    Drawing& d;
    GanttAxis axis;
    Emu th = 0, rowH = 0, msRowH = 0;
    Emu rowsTop = 0, rowsBottom = 0, msTop = 0;
    Box plot;
    Emu right = 0;  // rechter Rand des Diagramms
    std::wstring dateFormat;

    // Band einer Zeile: Aktivitäten ab rowsTop, Zeile 1 = Meilensteinzeile
    bool band(int row, Emu& y, Emu& h) const {
        if (row == kMilestoneRow) {
            y = msTop, h = msRowH;
            return true;
        }
        if (row < kFirstActivityRow || static_cast<size_t>(row) >= item.table.size()) return false;
        y = rowsTop + (row - kFirstActivityRow) * rowH, h = rowH;
        return true;
    }
    Emu x(double t) const { return static_cast<Emu>(std::llround(axis.xOf(t))); }
    bool visible(double s, double e) const { return e > axis.t0 && s < axis.t1; }

    // Automatische Platzierung der Beschriftungen einer Zeile: erst alle Formen der Zeile, dann die Beschriftungen
    // neben ihre Form (rechts bevorzugt), ohne Formen oder andere Beschriftungen zu überdecken.
    struct PendingLabel {
        Box anchor;        // Form, an der die Beschriftung steht
        std::wstring text;
        std::string role;
        Emu bandY = 0, bandH = 0;
    };
    std::vector<Box> taken;
    std::vector<PendingLabel> pending;

    static bool overlaps(const Box& a, const Box& b) { return a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h; }
    bool isFree(const Box& r) const {
        for (auto& t : taken)
            if (overlaps(r, t)) return false;
        return true;
    }

    void placeLabels() {
        const Emu gap = th / 6;
        for (auto& l : pending) {
            const Emu tw = textWidth(l.text, o.fontPt) + th / 4;
            const Emu ty = l.bandY + (l.bandH - th) / 2;
            std::optional<Box> chosen;
            // nach rechts: hinter der Form, sonst hinter dem, was im Weg steht
            for (Emu tx = l.anchor.x + l.anchor.w + gap; !chosen && tx + tw <= right;) {
                const Box r{tx, ty, tw, th};
                const auto hit = std::find_if(taken.begin(), taken.end(), [&](const Box& t) { return overlaps(r, t); });
                if (hit == taken.end()) chosen = r;
                else tx = hit->x + hit->w + gap;
            }
            // nach links
            for (Emu tx = l.anchor.x - gap - tw; !chosen && tx >= plot.x;) {
                const Box r{tx, ty, tw, th};
                const auto hit = std::find_if(taken.begin(), taken.end(), [&](const Box& t) { return overlaps(r, t); });
                if (hit == taken.end()) chosen = r;
                else tx = hit->x - gap - tw;
            }
            const bool leftOfAnchor = chosen && chosen->x < l.anchor.x;
            if (!chosen) chosen = Box{std::min(l.anchor.x + l.anchor.w + gap, right - tw), ty, tw, th};  // notfalls überdecken
            taken.push_back(*chosen);
            d.prims.push_back(text(chosen->x, l.bandY, tw, l.bandH, l.text, o.fontPt, leftOfAnchor ? Align::Right : Align::Left,
                                   Color::schemeColor(L"tx1"), l.role));
        }
        pending.clear();
        taken.clear();
    }

    // Balken (auch Fünfeck/Chevron) mit Pfeilen an abgeschnittenen Enden und Datumsbeschriftung
    void bar(const std::string& role, int row, double s, double e, const char* preset, const std::wstring& inner, Color fill, BarLabel label) {
        Emu y, h;
        if (!band(row, y, h) || !visible(s, e)) return;
        const Emu x0 = x(std::max(s, axis.t0)), x1 = x(std::min(e, axis.t1));
        const Emu barH = std::min<Emu>(h * 55 / 100, th);
        const Emu by = y + (h - barH) / 2;
        Prim p = preset ? shape(preset, x0, by, x1 - x0, barH, fill, role) : rect(x0, by, x1 - x0, barH, fill, role);
        p.flags = "bar";
        if (!inner.empty()) {
            p.text = inner;
            p.fontPt = o.fontPt;
            p.textColor = contrastOn(p.fill);
        }
        d.prims.push_back(p);
        taken.push_back({p.x, p.y, p.w, p.h});
        const Emu arrow = std::min<Emu>(barH * 6 / 10, std::max<Emu>((x1 - x0) / 3, 1));
        if (s < axis.t0)
            d.prims.push_back(shape("triangle", x0 + arrow / 4, by + barH / 4, arrow, barH / 2, contrastOn(p.fill), role + " from", 270));
        if (e > axis.t1)
            d.prims.push_back(shape("triangle", x1 - arrow - arrow / 4, by + barH / 4, arrow, barH / 2, contrastOn(p.fill), role + " to", 90));
        if (label != BarLabel::None) pending.push_back({{p.x, p.y, p.w, p.h}, rangeText(s, e, dateFormat, o.language, label), role + " label", y, h});
    }

    // Meilenstein (Raute) mit Datumsbeschriftung daneben
    void milestone(const std::string& role, const std::string& labelRole, int row, double t, Color fill, const std::wstring& caption) {
        Emu y, h;
        if (!band(row, y, h) || t < axis.t0 || t >= axis.t1) return;
        const double at = axis.dayScale ? std::floor(t) + 0.5 : std::floor(t);  // Tagesmitte nur mit Tagesskala
        const Emu size = std::min<Emu>(h * 65 / 100, th);
        const Emu cx = x(at);
        Prim p = shape("diamond", cx - size / 2, y + (h - size) / 2, size, size, fill, role);
        p.flags = "milestone";
        d.prims.push_back(p);
        taken.push_back({p.x, p.y, p.w, p.h});
        if (!caption.empty()) pending.push_back({{p.x, p.y, p.w, p.h}, caption, labelRole, y, h});
    }
};

}  // namespace

// ---- Zeitachse ----

double GanttAxis::xOf(double t) const {
    if (!hideWeekends) return static_cast<double>(x0) + (t - t0) / (t1 - t0) * static_cast<double>(w);
    const double total = workTime(t0, t1, weekendFirst, weekendLast);
    if (total <= 0) return static_cast<double>(x0);
    return static_cast<double>(x0) + workTime(t0, t, weekendFirst, weekendLast) / total * static_cast<double>(w);
}

double GanttAxis::dateAt(double xe) const {
    if (!hideWeekends) return t0 + (xe - static_cast<double>(x0)) / static_cast<double>(w) * (t1 - t0);
    double lo = t0 - 3660, hi = t1 + 3660;
    for (int i = 0; i < 64; ++i) {
        const double mid = (lo + hi) / 2;
        (xOf(mid) < xe ? lo : hi) = mid;
    }
    return hi;
}

double GanttAxis::emuPerDay() const {
    const double days = hideWeekends ? workTime(t0, t1, weekendFirst, weekendLast) : t1 - t0;
    return days > 0 ? static_cast<double>(w) / days : 0;
}

TimeUnit GanttAxis::snapUnit(double minEmu) const {
    const double perDay = emuPerDay();
    for (TimeUnit u : {TimeUnit::Day, TimeUnit::Week, TimeUnit::Month, TimeUnit::Quarter})
        if (perDay * unitDays(u) >= minEmu) return u;
    return TimeUnit::Year;
}

// ---- Daten ----

std::optional<double> cellDate(const Cell& cell) {
    if (cell.kind == CellKind::Date || cell.kind == CellKind::String) return parseIsoDate(cell.text);
    return std::nullopt;
}

std::optional<std::pair<double, double>> ganttDataRange(const DataItem& item, const ChartOptions& options) {
    std::optional<std::pair<double, double>> range;
    auto include = [&](double a, double b) {
        if (!range) range = {a, b};
        range->first = std::min(range->first, a);
        range->second = std::max(range->second, b);
    };
    for (size_t r = kMilestoneRow; r < item.table.size(); ++r) {
        const auto s = cellDate(at(item, r, kStartCol)), e = cellDate(at(item, r, kEndCol)), m = cellDate(at(item, r, kMilestoneCol));
        if (r >= kFirstActivityRow && s && e && *e > *s) include(*s, *e);
        if (m) include(std::floor(*m), std::floor(*m) + 1);
    }
    for (auto& it : options.gantt.items) {
        const auto s = parseIsoDate(it.start), e = parseIsoDate(it.end);
        if (it.kind == GanttItem::Kind::Milestone) {
            if (s) include(std::floor(*s), std::floor(*s) + 1);
        } else if (s && e && *e > *s) {
            include(*s, *e);
        }
    }
    return range;
}

std::pair<double, double> ganttRange(const DataItem& item, const ChartOptions& options) {
    const GanttOptions& g = options.gantt;
    std::optional<double> s = parseIsoDate(g.start), e = parseIsoDate(g.end);
    if (s) s = std::floor(*s);
    if (e) e = std::floor(*e) + 1;  // letzter Tag eingeschlossen
    if (!s || !e) {
        double a = 0, b = 0;
        if (const auto data = ganttDataRange(item, options)) {
            a = static_cast<double>(unitStart(data->first, TimeUnit::Month, g.weekStart));
            b = static_cast<double>(nextUnit(unitStart(data->second - 1e-6, TimeUnit::Month, g.weekStart), TimeUnit::Month));
        } else {
            a = static_cast<double>(unitStart(g.referenceDay, TimeUnit::Month, g.weekStart));
            b = addUnits(a, TimeUnit::Month, 3);
        }
        if (!s) s = a;
        if (!e) e = std::max(b, *s + 1);
    }
    if (*e <= *s) e = *s + 1;
    return {*s, *e};
}

std::vector<TimeUnit> ganttScales(double t0, double t1, Emu plotWidth, const ChartOptions& options) {
    if (!options.gantt.scales.empty()) {
        std::vector<TimeUnit> s = options.gantt.scales;
        std::sort(s.begin(), s.end(), [](TimeUnit a, TimeUnit b) { return static_cast<int>(a) > static_cast<int>(b); });
        s.erase(std::unique(s.begin(), s.end()), s.end());
        return s;
    }
    const double perDay = static_cast<double>(plotWidth) / std::max(t1 - t0, 1.0);
    TimeUnit finest = TimeUnit::Year;
    for (TimeUnit u : {TimeUnit::Day, TimeUnit::Week, TimeUnit::Month, TimeUnit::Quarter, TimeUnit::Year}) {
        if (!scaleFormat(u, perDay * minUnitDays(u), options).empty()) {
            finest = u;
            break;
        }
    }
    if (finest == TimeUnit::Year) return {TimeUnit::Year};
    return {parentUnit(finest), finest};
}

GanttAxis ganttAxis(const DataItem& item, const ChartOptions& options, Box bounds) { return layoutChart(item, options, bounds).axis; }

Drawing layoutGantt(const DataItem& item, const ChartOptions& o, Box b) {
    Drawing d;
    d.bounds = b;
    d.valueDirection = Orientation::Right;
    const GanttOptions& g = o.gantt;
    if (item.table.empty()) throw std::invalid_argument("\"table\" ist leer");
    for (size_t r = kMilestoneRow; r < item.table.size(); ++r) {
        const auto s = cellDate(at(item, r, kStartCol)), e = cellDate(at(item, r, kEndCol));
        if (r >= kFirstActivityRow && s && e && *e <= *s)
            d.warnings.push_back(std::format("Zeile {}: Ende liegt nicht nach dem Start – kein Balken", r + 1));
        for (int c : {kStartCol, kEndCol, kMilestoneCol}) {
            const Cell& cell = at(item, r, c);
            if (cell.kind != CellKind::Empty && !cellDate(cell))
                d.warnings.push_back(std::format("Zeile {}, Spalte {}: kein Datum – ignoriert", r + 1, c + 1));
        }
    }

    Layout L{item, o, d};
    L.th = textHeight(o.fontPt);
    const Emu th = L.th;
    L.dateFormat = g.dateFormat.empty() ? L"dd.MM." : g.dateFormat;

    // Beschriftungsspalten
    const int rows = std::max(0, static_cast<int>(item.table.size()) - kFirstActivityRow);
    auto hasText = [&](int col) {
        for (size_t r = kFirstActivityRow; r < item.table.size(); ++r)
            if (!textAt(item, r, col).empty()) return true;
        return false;
    };
    struct Column {
        int col;
        std::string key;
        std::wstring head;
        Emu w = 0, x = 0;
    };
    std::vector<Column> columns{{kLabelCol, "activity", headline(item, kLabelCol, L"Aktivität", L"Activity", o.language)}};
    if (g.responsible.value_or(hasText(kResponsibleCol)))
        columns.push_back({kResponsibleCol, "responsible", headline(item, kResponsibleCol, L"Verantwortlich", L"Responsible", o.language)});
    if (g.remark.value_or(hasText(kRemarkCol))) columns.push_back({kRemarkCol, "remark", headline(item, kRemarkCol, L"Bemerkung", L"Remark", o.language)});
    Emu labelsW = 0;
    for (auto& c : columns) {
        c.w = g.headlines ? textWidth(c.head, o.fontPt) : 0;
        for (size_t r = kMilestoneRow; r < item.table.size(); ++r) c.w = std::max(c.w, textWidth(textAt(item, r, c.col), o.fontPt));
        c.w = std::max(c.w + th / 2, th * 2);
        labelsW += c.w;
    }
    if (labelsW > b.w / 2) {  // höchstens die halbe Breite
        for (auto& c : columns) c.w = c.w * (b.w / 2) / labelsW;
        labelsW = b.w / 2;
    }
    Emu cx = b.x;
    for (auto& c : columns) c.x = cx, cx += c.w;

    L.plot = {b.x + labelsW, b.y, b.w - labelsW - th / 3, b.h};
    L.right = b.x + b.w;
    if (L.plot.w < th * 4) throw std::invalid_argument("Platzhalter ist zu schmal für den Zeitplan");

    // Zeitachse und Skalen
    const auto [t0, t1] = ganttRange(item, o);
    GanttAxis& a = L.axis;
    a.valid = true;
    a.t0 = t0, a.t1 = t1, a.x0 = L.plot.x, a.w = L.plot.w;
    a.hideWeekends = g.weekends == WeekendMode::Hidden;
    a.weekendFirst = g.weekendFirst, a.weekendLast = g.weekendLast, a.weekStart = g.weekStart;
    const double effectiveDays = a.hideWeekends ? std::max(workTime(t0, t1, a.weekendFirst, a.weekendLast), 1.0) : t1 - t0;
    const auto scales = ganttScales(0, effectiveDays, L.plot.w, o);
    a.finest = scales.back();
    a.dayScale = std::find(scales.begin(), scales.end(), TimeUnit::Day) != scales.end();

    const Emu scaleH = th + th / 2;
    const Emu scalesH = scaleH * static_cast<Emu>(scales.size());
    L.msRowH = th + th * 2 / 3;
    L.rowsTop = b.y + scalesH;
    // Beschriftungen der Schattierungen stehen in einer eigenen Zeile unter dem Diagramm
    const bool shadeCaptions = std::any_of(g.items.begin(), g.items.end(), [&](const GanttItem& it) {
        const auto s = parseIsoDate(it.start), e = parseIsoDate(it.end);
        return it.kind == GanttItem::Kind::Shade && s && e && *e > *s && *e > t0 && *s < t1;
    });
    const Emu shadeCaptionH = shadeCaptions ? th + th / 6 : 0;
    const Emu available = b.h - scalesH - L.msRowH - shadeCaptionH;
    L.rowH = rows ? std::min<Emu>(available / rows, th * 5 / 2) : 0;
    if (rows && L.rowH < th * 11 / 10) throw std::invalid_argument("Platzhalter ist zu niedrig für alle Zeilen des Zeitplans");
    L.rowsBottom = L.rowsTop + L.rowH * rows;
    L.msTop = L.rowsBottom;
    const Emu chartRight = b.x + b.w;

    // 1. Hintergrund: Zeilenschattierung, Wochenenden, Schattierungen
    for (int r : g.shadedRows) {
        Emu y, h;
        if (r >= kFirstActivityRow && L.band(r, y, h)) d.prims.push_back(rect(b.x, y, b.w, h, Color::rgbColor(kRowShade), std::format("row shade r{}", r)));
    }
    // Wochenenden nur schattieren, wenn Tage erkennbar sind (Tagesskala oder breite Tage)
    if (g.weekends == WeekendMode::Shaded && (a.dayScale || a.emuPerDay() >= th * 0.6) && rows > 0) {
        int k = 0;
        for (long long day = static_cast<long long>(std::floor(t0)); static_cast<double>(day) < t1;) {
            if (!isWeekend(day, a.weekendFirst, a.weekendLast)) {
                ++day;
                continue;
            }
            long long end = day;
            while (static_cast<double>(end) < t1 && isWeekend(end, a.weekendFirst, a.weekendLast)) ++end;
            const Emu x0 = L.x(std::max(static_cast<double>(day), t0)), x1 = L.x(std::min(static_cast<double>(end), t1));
            d.prims.push_back(rect(x0, L.rowsTop, x1 - x0, L.rowsBottom - L.rowsTop, Color::rgbColor(kWeekend), std::format("weekend {}", ++k)));
            day = end;
        }
    }
    for (size_t i = 0; i < g.items.size(); ++i) {
        const GanttItem& it = g.items[i];
        if (it.kind != GanttItem::Kind::Shade) continue;
        const auto s = parseIsoDate(it.start), e = parseIsoDate(it.end);
        if (!s || !e || *e <= *s || !L.visible(*s, *e)) continue;
        const Emu x0 = L.x(std::max(*s, t0)), x1 = L.x(std::min(*e, t1));
        const Emu bottom = L.msTop + L.msRowH;
        Prim p = rect(x0, L.rowsTop, x1 - x0, std::max<Emu>(bottom - L.rowsTop, 1), Color::rgbColor(it.fill.value_or(kShade)), std::format("item {}", i));
        p.flags = "shade";
        d.prims.push_back(p);
        const std::wstring caption = it.text.empty() ? rangeText(*s, *e, L.dateFormat, o.language, BarLabel::Range) : it.text;
        const Emu tw = std::max<Emu>(x1 - x0, textWidth(caption, o.fontPt) + th / 4);
        d.prims.push_back(text(std::clamp<Emu>((x0 + x1 - tw) / 2, L.plot.x, L.right - tw), bottom + th / 6, tw, th, caption, o.fontPt, Align::Center,
                               Color::schemeColor(L"tx1"), std::format("item {} label", i)));
    }

    // 2. Trennlinien an den Skaleneinheiten (grobe durchgezogen, feine gestrichelt)
    std::vector<long long> taken;
    if (g.separators && rows > 0) {
        for (size_t k = 0; k < scales.size(); ++k) {
            const TimeUnit u = scales[k];
            if (a.emuPerDay() * minUnitDays(u) < th * 0.6) continue;  // zu dicht
            int n = 0;
            for (long long s = nextUnit(unitStart(t0, u, g.weekStart), u); static_cast<double>(s) < t1; s = nextUnit(s, u)) {
                if (static_cast<double>(s) <= t0 || std::find(taken.begin(), taken.end(), s) != taken.end()) continue;
                taken.push_back(s);
                Prim p = line(L.x(static_cast<double>(s)), L.rowsTop, L.x(static_cast<double>(s)), L.rowsBottom, Color::rgbColor(kGrid), 0.5,
                              std::format("separator {} {}", unitKey(u), ++n));
                p.dashed = k > 0;
                d.prims.push_back(p);
            }
        }
    }
    const auto today = g.todayLine ? (g.today.empty() ? std::optional<double>(g.referenceDay) : parseIsoDate(g.today)) : std::nullopt;
    if (today && *today >= t0 && *today < t1) {
        const Emu tx = L.x(std::floor(*today) + 0.5);
        Prim p = line(tx, L.rowsTop, tx, L.msTop + L.msRowH, Color::rgbColor(kToday), 1.0, "today");
        p.dashed = true;
        d.prims.push_back(p);
    }

    // 3. Skalen: Beschriftungen je Einheit, Teilstriche, Linien unter jeder Skala
    for (size_t k = 0; k < scales.size(); ++k) {
        const TimeUnit u = scales[k];
        const Emu y = b.y + scaleH * static_cast<Emu>(k);
        const std::wstring fmt = scaleFormat(u, a.emuPerDay() * minUnitDays(u), o);
        int n = 0;
        for (long long s = unitStart(t0, u, g.weekStart); static_cast<double>(s) < t1; s = nextUnit(s, u)) {
            ++n;
            const double from = std::max(static_cast<double>(s), t0), to = std::min(static_cast<double>(nextUnit(s, u)), t1);
            const Emu x0 = L.x(from), x1 = L.x(to);
            if (static_cast<double>(s) > t0)
                d.prims.push_back(line(x0, y + scaleH / 3, x0, y + scaleH, Color::rgbColor(kGrid), 0.5, std::format("scale {} tick {}", unitKey(u), n)));
            if (fmt.empty()) continue;
            const std::wstring label = formatDate(static_cast<double>(s), fmt, o.language);
            if (x1 - x0 < textWidth(label, o.fontPt) + th / 4) continue;  // angeschnittene Einheit am Rand
            d.prims.push_back(text(x0, y, x1 - x0, scaleH, label, o.fontPt, Align::Center, Color::schemeColor(L"tx1"),
                                   std::format("scale {} {}", unitKey(u), n)));
        }
        const bool last = k + 1 == scales.size();
        d.prims.push_back(line(L.plot.x, y + scaleH, L.plot.x + L.plot.w, y + scaleH, last ? Color::schemeColor(L"tx1") : Color::rgbColor(kGrid),
                               last ? 0.75 : 0.5, std::format("scale {} rule", unitKey(u))));
    }
    d.prims.push_back(line(L.plot.x, b.y, L.plot.x, L.rowsTop, Color::rgbColor(kGrid), 0.5, "scale start"));
    d.prims.push_back(line(L.plot.x + L.plot.w, b.y, L.plot.x + L.plot.w, L.rowsTop, Color::rgbColor(kGrid), 0.5, "scale end"));

    // 4. Überschriften der Beschriftungsspalten (Höhe der feinsten Skala)
    if (g.headlines)
        for (auto& c : columns) {
            Prim p = text(c.x + th / 4, L.rowsTop - scaleH, c.w - th / 4, scaleH, c.head, o.fontPt, Align::Left, Color::schemeColor(L"tx1"),
                          "headline " + c.key);
            p.bold = true;
            d.prims.push_back(p);
        }

    // 5. Zeilen: erst alle Formen einer Zeile (Datenblatt und frei gesetzte), dann ihre Beschriftungen
    auto rowContent = [&](int r) {
        Emu y, h;
        L.band(r, y, h);
        // Klammern oben in der Zeile: Beschriftung, darunter die Linie mit Enden nach unten
        for (size_t i = 0; i < g.items.size(); ++i) {
            const GanttItem& it = g.items[i];
            if (it.row != r || it.kind != GanttItem::Kind::Bracket) continue;
            const auto s = parseIsoDate(it.start), e = parseIsoDate(it.end);
            if (!s || !e || *e <= *s || !L.visible(*s, *e)) continue;
            const std::string role = std::format("item {}", i);
            const Emu x0 = L.x(std::max(*s, t0)), x1 = L.x(std::min(*e, t1));
            const Emu ly = y + th + th / 8, tick = th / 3;
            Prim p = line(x0, ly, x1, ly, Color::schemeColor(L"tx1"), 1.0, role);
            p.flags = "bracket";
            d.prims.push_back(p);
            d.prims.push_back(line(x0, ly, x0, ly + tick, Color::schemeColor(L"tx1"), 1.0, role + " left"));
            d.prims.push_back(line(x1, ly, x1, ly + tick, Color::schemeColor(L"tx1"), 1.0, role + " right"));
            const std::wstring caption = it.text.empty() ? rangeText(*s, *e, L.dateFormat, o.language, BarLabel::Range) : it.text;
            const Emu tw = std::max<Emu>(x1 - x0, textWidth(caption, o.fontPt) + th / 4);
            d.prims.push_back(text((x0 + x1 - tw) / 2, y, tw, th, caption, o.fontPt, Align::Center, Color::schemeColor(L"tx1"), role + " label"));
            L.taken.push_back({std::min(x0, (x0 + x1 - tw) / 2), y, std::max(x1 - x0, tw), ly + tick - y});
        }
        // Balken und Meilenstein aus dem Datenblatt
        if (r >= kFirstActivityRow) {
            const Cell& sc = at(item, r, kStartCol);
            const auto s = cellDate(sc), e = cellDate(at(item, r, kEndCol));
            if (s && e && *e > *s)
                L.bar(std::format("bar r{}", r), r, *s, *e, nullptr, L"", sc.fill ? Color::rgbColor(*sc.fill) : Color::schemeColor(L"accent1"),
                      g.barLabels);
        }
        // Frei gesetzte Balken, Fünfecke, Chevrons
        for (size_t i = 0; i < g.items.size(); ++i) {
            const GanttItem& it = g.items[i];
            if (it.row != r || (it.kind != GanttItem::Kind::Bar && it.kind != GanttItem::Kind::Pentagon && it.kind != GanttItem::Kind::Chevron)) continue;
            if (r == kMilestoneRow) {
                d.warnings.push_back(std::format("Element {}: die Meilensteinzeile nimmt keine Balken auf", i + 1));
                continue;
            }
            const auto s = parseIsoDate(it.start), e = parseIsoDate(it.end);
            if (!s || !e || *e <= *s) continue;
            const char* preset = it.kind == GanttItem::Kind::Pentagon ? "homePlate" : it.kind == GanttItem::Kind::Chevron ? "chevron" : nullptr;
            L.bar(std::format("item {}", i), r, *s, *e, preset, it.kind == GanttItem::Kind::Bar ? L"" : it.text,
                  it.fill ? Color::rgbColor(*it.fill) : Color::schemeColor(L"accent1"), it.kind == GanttItem::Kind::Bar ? g.barLabels : BarLabel::None);
        }
        // Meilensteine (Datenblatt, dann frei gesetzte)
        const Cell& mc = at(item, r, kMilestoneCol);
        if (const auto m = cellDate(mc))
            L.milestone(std::format("milestone r{}", r), std::format("milestone label r{}", r), r, *m,
                        mc.fill ? Color::rgbColor(*mc.fill) : Color::schemeColor(L"accent2"),
                        g.milestoneLabels ? formatDate(*m, L.dateFormat, o.language) : L"");
        for (size_t i = 0; i < g.items.size(); ++i) {
            const GanttItem& it = g.items[i];
            if (it.row != r || it.kind != GanttItem::Kind::Milestone) continue;
            if (const auto s = parseIsoDate(it.start)) {
                const std::string role = std::format("item {}", i);
                L.milestone(role, role + " label", r, *s, it.fill ? Color::rgbColor(*it.fill) : Color::schemeColor(L"accent2"),
                            it.text.empty() ? (g.milestoneLabels ? formatDate(*s, L.dateFormat, o.language) : L"") : it.text);
            }
        }
        L.placeLabels();
    };

    for (int r = kFirstActivityRow; r < static_cast<int>(item.table.size()); ++r) {
        Emu y, h;
        L.band(r, y, h);
        if (g.rowSeparators && r > kFirstActivityRow)
            d.prims.push_back(line(b.x, y, chartRight, y, Color::rgbColor(kGrid), 0.5, std::format("row line r{}", r)));
        for (auto& c : columns) {
            const std::wstring s = textAt(item, r, c.col);
            if (!s.empty())
                d.prims.push_back(text(c.x + th / 4, y, c.w - th / 4, h, s, o.fontPt, Align::Left, Color::schemeColor(L"tx1"),
                                       std::format("{} r{}", c.key, r)));
        }
        rowContent(r);
    }

    // 6. Meilensteinzeile unter dem Diagramm (nimmt auch Meilensteine und Klammern auf)
    d.prims.push_back(line(L.plot.x, L.msTop, L.plot.x + L.plot.w, L.msTop, Color::schemeColor(L"tx1"), 0.75, "milestone row rule"));
    if (const std::wstring caption = textAt(item, kMilestoneRow, kLabelCol); !caption.empty())
        d.prims.push_back(text(columns[0].x + th / 4, L.msTop, columns[0].w - th / 4, L.msRowH, caption, o.fontPt, Align::Left,
                               Color::schemeColor(L"tx1"), "activity r1"));
    if (const auto m = cellDate(at(item, kMilestoneRow, kMilestoneCol)); m && g.milestoneLines && *m >= t0 && *m < t1 && rows > 0) {
        const Emu mx = L.x(a.dayScale ? std::floor(*m) + 0.5 : std::floor(*m));
        Prim p = line(mx, L.rowsTop, mx, L.msTop, Color::schemeColor(L"accent2"), 0.75, "milestone line r1");
        p.dashed = true;
        d.prims.push_back(p);
    }
    rowContent(kMilestoneRow);

    for (size_t i = 0; i < g.items.size(); ++i) {
        const int r = g.items[i].row;
        if (g.items[i].kind != GanttItem::Kind::Shade && r != kMilestoneRow && (r < kFirstActivityRow || r >= static_cast<int>(item.table.size())))
            d.warnings.push_back(std::format("Element {}: Zeile {} gibt es nicht", i + 1, r + 1));
    }

    for (int r = kFirstActivityRow; r < static_cast<int>(item.table.size()); ++r) d.rows.push_back({r, L.rowsTop + (r - kFirstActivityRow) * L.rowH, L.rowH});
    d.rows.push_back({kMilestoneRow, L.msTop, L.msRowH});
    d.scaleTop = b.y, d.scaleBottom = L.rowsTop;
    d.axis = a;
    return d;
}

}  // namespace rethink::engine
