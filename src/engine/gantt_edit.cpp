#include "gantt_edit.h"

#include "dates.h"

#include <algorithm>
#include <charconv>
#include <cmath>
#include <format>
#include <stdexcept>
#include <string_view>

namespace rethink::engine {
namespace {

using namespace gantt;

Cell& cellAt(DataItem& item, int r, int c) {
    if (static_cast<int>(item.table.size()) <= r) item.table.resize(r + 1);
    Row& row = item.table[r];
    if (static_cast<int>(row.size()) <= c) row.resize(c + 1);
    return row[c];
}

const Cell* cellIf(const DataItem& item, int r, int c) {
    if (r < 0 || r >= static_cast<int>(item.table.size()) || c < 0 || c >= static_cast<int>(item.table[r].size())) return nullptr;
    return &item.table[r][c];
}

void setDate(DataItem& item, int r, int c, double t) {
    Cell& cell = cellAt(item, r, c);
    cell.kind = CellKind::Date;
    cell.text = isoDate(t);
    cell.number = 0;
}

void clearCell(DataItem& item, int r, int c) {
    if (Cell* cell = const_cast<Cell*>(cellIf(item, r, c))) *cell = Cell{};
}

// Elementzeilen ab `from` um `delta` verschieben (Einfügen/Löschen von Zeilen).
void shiftRows(ChartOptions& options, int from, int delta) {
    for (auto& it : options.gantt.items)
        if (it.row >= from) it.row += delta;
    for (auto& r : options.gantt.shadedRows)
        if (r >= from) r += delta;
    for (auto& s : options.gantt.rowStyles)
        if (s.row >= from) s.row += delta;
}

// "bar r12" mit Präfix "bar r" → 12; genau Präfix + Zahl.
bool roleNumber(const std::string& role, std::string_view prefix, int& n) {
    if (!role.starts_with(prefix) || role.size() == prefix.size()) return false;
    const auto digits = std::string_view(role).substr(prefix.size());
    const auto [end, error] = std::from_chars(digits.data(), digits.data() + digits.size(), n);
    return error == std::errc() && end == digits.data() + digits.size();
}

}  // namespace

GanttTarget ganttTarget(const DataItem& item, const ChartOptions& options, const std::string& role) {
    GanttTarget t;
    int n = -1;
    if (roleNumber(role, "bar r", n)) {
        const Cell *s = cellIf(item, n, kStartCol), *e = cellIf(item, n, kEndCol);
        const auto sd = s ? cellDate(*s) : std::nullopt, ed = e ? cellDate(*e) : std::nullopt;
        if (!sd || !ed) return t;
        t.kind = GanttTarget::Kind::Bar, t.row = n, t.start = *sd, t.end = *ed;
    } else if (roleNumber(role, "milestone r", n)) {
        const Cell* m = cellIf(item, n, kMilestoneCol);
        const auto md = m ? cellDate(*m) : std::nullopt;
        if (!md) return t;
        t.kind = GanttTarget::Kind::Milestone, t.row = n, t.start = t.end = *md;
    } else if (roleNumber(role, "item ", n) && n >= 0 && n < static_cast<int>(options.gantt.items.size())) {
        const GanttItem& it = options.gantt.items[n];
        const auto s = parseIsoDate(it.start), e = parseIsoDate(it.end);
        if (!s) return t;
        t.item = n, t.row = it.row, t.start = *s, t.end = e.value_or(*s);
        switch (it.kind) {
            case GanttItem::Kind::Milestone: t.kind = GanttTarget::Kind::Milestone; break;
            case GanttItem::Kind::Bracket: t.kind = GanttTarget::Kind::Bracket; break;
            case GanttItem::Kind::Shade: t.kind = GanttTarget::Kind::Shade; break;
            default: t.kind = GanttTarget::Kind::Bar; break;
        }
        if (t.kind != GanttTarget::Kind::Milestone && !e) t.kind = GanttTarget::Kind::None;
    }
    return t;
}

void setGanttDates(DataItem& item, ChartOptions& options, const GanttTarget& target, double start, double end) {
    if (target.kind == GanttTarget::Kind::None) return;
    if (target.item >= 0) {
        GanttItem& it = options.gantt.items.at(target.item);
        it.start = isoDate(start);
        if (target.kind != GanttTarget::Kind::Milestone) it.end = isoDate(end);
        return;
    }
    if (target.kind == GanttTarget::Kind::Milestone) {
        setDate(item, target.row, kMilestoneCol, start);
    } else {
        setDate(item, target.row, kStartCol, start);
        setDate(item, target.row, kEndCol, end);
    }
}

std::pair<double, double> ganttDragDates(const GanttAxis& axis, const GanttTarget& target, GanttGrip grip, double dx, TimeUnit unit) {
    const int ws = axis.weekStart;
    if (target.kind == GanttTarget::Kind::Milestone) {
        // Meilensteine stehen in der Tagesmitte (Tagesskala) bzw. am Tagesanfang: der Tag unter der Maus zählt
        const double shown = axis.dayScale ? std::floor(target.start) + 0.5 : std::floor(target.start);
        const double at = axis.dateAt(axis.xOf(shown) + dx);
        const double day = unit == TimeUnit::Day ? std::floor(at) : snapToUnit(at, unit, ws);
        return {day, day};
    }
    const double s = target.start, e = target.end;
    switch (grip) {
        case GanttGrip::Start: {
            double t = snapToUnit(axis.dateAt(axis.xOf(s) + dx), unit, ws);
            if (t > e - 1) t = e - 1;  // mindestens ein Tag
            return {t, e};
        }
        case GanttGrip::End: {
            double t = snapToUnit(axis.dateAt(axis.xOf(e) + dx), unit, ws);
            if (t < s + 1) t = s + 1;
            return {s, t};
        }
        case GanttGrip::Move: {
            const double moved = snapToUnit(axis.dateAt(axis.xOf(s) + dx), unit, ws);
            const double shift = std::round(moved - s);  // ganze Tage: die Dauer bleibt erhalten
            return {s + shift, e + shift};
        }
    }
    return {s, e};
}

std::pair<double, double> ganttRangeDrag(const GanttAxis& axis, bool startSide, double dx, TimeUnit unit) {
    const int ws = axis.weekStart;
    double t0 = axis.t0, t1 = axis.t1;
    if (startSide) {
        t0 = snapToUnit(axis.dateAt(static_cast<double>(axis.x0) + dx), unit, ws);
        if (t0 >= t1) t0 = static_cast<double>(unitStart(t1 - 1e-6, unit, ws));
        if (t0 >= t1) t0 = t1 - 1;
    } else {
        t1 = snapToUnit(axis.dateAt(static_cast<double>(axis.x0 + axis.w) + dx), unit, ws);
        if (t1 <= t0) t1 = static_cast<double>(nextUnit(unitStart(t0, unit, ws), unit));
        if (t1 <= t0) t1 = t0 + 1;
    }
    return {t0, t1};
}

std::wstring ganttTooltip(double t, const std::wstring& language) {
    // Datum, dahinter Wochentag und Kalenderwoche: „04.11.2026 (Mi der KW 45)“
    return std::format(L"{} ({} {} {})", formatDate(t, language == L"en" ? L"M/d/yyyy" : L"dd.MM.yyyy", language), formatDate(t, L"ddd", language),
                       language == L"en" ? L"of" : L"der", formatDate(t, L"www", language));
}

void setGanttRange(ChartOptions& options, double t0, double t1) {
    options.gantt.start = isoDate(std::floor(t0));
    options.gantt.end = isoDate(std::ceil(t1 - 1e-9) - 1);  // letzter sichtbarer Tag
}

bool fitGanttRange(const DataItem& item, ChartOptions& options) {
    const auto range = ganttDataRange(item, options);
    if (!range) return false;
    setGanttRange(options, std::floor(range->first), std::ceil(range->second - 1e-9));
    return true;
}

int insertGanttRow(DataItem& item, ChartOptions& options, int before, const std::wstring& label) {
    while (static_cast<int>(item.table.size()) < kFirstActivityRow) item.table.push_back(Row{});
    before = std::clamp(before, kFirstActivityRow, static_cast<int>(item.table.size()));
    Row row(1);
    row[0] = Cell{CellKind::String, label, 0, std::nullopt};
    item.table.insert(item.table.begin() + before, std::move(row));
    shiftRows(options, before, 1);
    return before;
}

void duplicateGanttRow(DataItem& item, ChartOptions& options, int row, bool above) {
    if (row < kFirstActivityRow || row >= static_cast<int>(item.table.size())) throw std::invalid_argument("Diese Zeile gibt es nicht");
    const Row copy = item.table[row];
    const int at = above ? row : row + 1;
    item.table.insert(item.table.begin() + at, copy);
    // Frei gesetzte Elemente der Zeile mitkopieren
    std::vector<GanttItem> added;
    for (auto& it : options.gantt.items)
        if (it.row == row && it.kind != GanttItem::Kind::Shade) added.push_back(it);
    std::optional<RowStyle> style;
    for (auto& s : options.gantt.rowStyles)
        if (s.row == row) style = s;
    shiftRows(options, at, 1);
    for (auto& it : added) {
        it.row = at;
        options.gantt.items.push_back(it);
    }
    if (style) {
        style->row = at;
        options.gantt.rowStyles.push_back(*style);
    }
}

void deleteGanttRow(DataItem& item, ChartOptions& options, int row) {
    if (row < kFirstActivityRow || row >= static_cast<int>(item.table.size())) throw std::invalid_argument("Diese Zeile gibt es nicht");
    item.table.erase(item.table.begin() + row);
    auto& items = options.gantt.items;
    items.erase(std::remove_if(items.begin(), items.end(), [&](const GanttItem& it) { return it.row == row && it.kind != GanttItem::Kind::Shade; }),
                items.end());
    auto& shaded = options.gantt.shadedRows;
    shaded.erase(std::remove(shaded.begin(), shaded.end(), row), shaded.end());
    auto& styles = options.gantt.rowStyles;
    styles.erase(std::remove_if(styles.begin(), styles.end(), [&](const RowStyle& s) { return s.row == row; }), styles.end());
    shiftRows(options, row + 1, -1);
}

bool moveGanttRow(DataItem& item, ChartOptions& options, int row, int delta) {
    const int target = row + delta;
    if (row < kFirstActivityRow || target < kFirstActivityRow || row >= static_cast<int>(item.table.size()) ||
        target >= static_cast<int>(item.table.size()))
        return false;
    std::swap(item.table[row], item.table[target]);
    for (auto& it : options.gantt.items) {
        if (it.row == row) it.row = target;
        else if (it.row == target) it.row = row;
    }
    for (auto& r : options.gantt.shadedRows) {
        if (r == row) r = target;
        else if (r == target) r = row;
    }
    for (auto& s : options.gantt.rowStyles) {
        if (s.row == row) s.row = target;
        else if (s.row == target) s.row = row;
    }
    return true;
}

bool moveGanttRowTo(DataItem& item, ChartOptions& options, int row, int target) {
    const int count = static_cast<int>(item.table.size());
    if (row < kFirstActivityRow || row >= count || target < kFirstActivityRow || target >= count) return false;
    const bool categories = !ganttCategories(item).empty();
    const Cell* targetCategory = cellIf(item, target, kCategoryCol);
    const std::wstring category = targetCategory && targetCategory->kind == CellKind::String ? targetCategory->text : L"";
    const Cell* ownCategory = cellIf(item, row, kCategoryCol);
    const bool sameCategory = (ownCategory && ownCategory->kind == CellKind::String ? ownCategory->text : L"") == category;
    if (row == target && (!categories || sameCategory)) return false;
    // neue Lage jeder Zeile: die gezogene an `target`, die dazwischen rücken um eins
    std::vector<int> newIndex(count);
    for (int r = 0; r < count; ++r) {
        if (r == row) newIndex[r] = target;
        else if (row < target && r > row && r <= target) newIndex[r] = r - 1;
        else if (row > target && r >= target && r < row) newIndex[r] = r + 1;
        else newIndex[r] = r;
    }
    std::vector<Row> moved(count);
    for (int r = 0; r < count; ++r) moved[newIndex[r]] = std::move(item.table[r]);
    item.table = std::move(moved);
    auto remap = [&](int& r) {
        if (r >= 0 && r < count) r = newIndex[r];
    };
    for (auto& it : options.gantt.items) remap(it.row);
    for (auto& r : options.gantt.shadedRows) remap(r);
    for (auto& s : options.gantt.rowStyles) remap(s.row);
    if (categories) setGanttText(item, target, kCategoryCol, category);  // in die Kategorie der Ablagestelle
    return true;
}

void addGanttItem(DataItem& item, ChartOptions& options, GanttItem::Kind kind, int row, double date, double days) {
    const double day = std::floor(date);
    const bool activity = row >= kFirstActivityRow && row < static_cast<int>(item.table.size());
    if (kind == GanttItem::Kind::Bar && activity) {
        const Cell *s = cellIf(item, row, kStartCol), *e = cellIf(item, row, kEndCol);
        if ((!s || !cellDate(*s)) && (!e || !cellDate(*e))) {
            setDate(item, row, kStartCol, day);
            setDate(item, row, kEndCol, day + days);
            return;
        }
    }
    if (kind == GanttItem::Kind::Milestone && (activity || row == kMilestoneRow)) {
        const Cell* m = cellIf(item, row, kMilestoneCol);
        if (!m || !cellDate(*m)) {
            setDate(item, row, kMilestoneCol, day);
            return;
        }
    }
    if (kind != GanttItem::Kind::Shade && kind != GanttItem::Kind::Milestone && kind != GanttItem::Kind::Bracket && !activity)
        throw std::invalid_argument("Balken gehören in eine Aktivitätszeile");
    GanttItem it;
    it.kind = kind;
    it.row = row;
    it.start = isoDate(day);
    if (kind != GanttItem::Kind::Milestone) it.end = isoDate(day + days);
    options.gantt.items.push_back(it);
}

void setGanttText(DataItem& item, int row, int column, const std::wstring& text) {
    if (text.empty()) {
        if (Cell* cell = const_cast<Cell*>(cellIf(item, row, column))) {
            const auto fill = cell->fill;
            *cell = Cell{};
            cell->fill = fill;
        }
        return;
    }
    Cell& cell = cellAt(item, row, column);
    cell.kind = CellKind::String;
    cell.text = text;
}

namespace {

std::vector<std::wstring> splitNames(const std::wstring& list) {
    std::vector<std::wstring> names;
    for (size_t pos = 0; pos <= list.size();) {
        const size_t sep = std::min(list.find_first_of(L";,", pos), list.size());
        std::wstring name = list.substr(pos, sep - pos);
        const auto a = name.find_first_not_of(L" \t"), b = name.find_last_not_of(L" \t");
        if (a != std::wstring::npos) names.push_back(name.substr(a, b - a + 1));
        pos = sep + 1;
    }
    return names;
}

}  // namespace

bool toggleGanttLink(DataItem& item, int row, const std::wstring& predecessor) {
    const Cell* cell = cellIf(item, row, kPredecessorCol);
    std::vector<std::wstring> names = splitNames(cell && cell->kind == CellKind::String ? cell->text : L"");
    const auto it = std::find(names.begin(), names.end(), predecessor);
    const bool linked = it == names.end();
    if (linked) names.push_back(predecessor);
    else names.erase(it);
    std::wstring joined;
    for (auto& n : names) joined += (joined.empty() ? L"" : L"; ") + n;
    setGanttText(item, row, kPredecessorCol, joined);
    return linked;
}

void sortGanttByCategory(DataItem& item, ChartOptions& options, const std::vector<std::wstring>& order) {
    const int first = kFirstActivityRow, count = static_cast<int>(item.table.size()) - first;
    if (count <= 1) return;
    std::vector<std::wstring> ranking = order;
    for (auto& c : ganttCategories(item))
        if (std::find(ranking.begin(), ranking.end(), c) == ranking.end()) ranking.push_back(c);
    auto rank = [&](int r) {
        const Cell* cell = cellIf(item, r, kCategoryCol);
        std::wstring c = cell && cell->kind == CellKind::String ? cell->text : L"";
        const auto a = c.find_first_not_of(L" \t"), b = c.find_last_not_of(L" \t");
        c = a == std::wstring::npos ? L"" : c.substr(a, b - a + 1);
        const auto it = std::find(ranking.begin(), ranking.end(), c);
        return c.empty() || it == ranking.end() ? static_cast<int>(ranking.size()) : static_cast<int>(it - ranking.begin());
    };
    std::vector<int> rows(count);
    for (int i = 0; i < count; ++i) rows[i] = first + i;
    std::stable_sort(rows.begin(), rows.end(), [&](int a, int b) { return rank(a) < rank(b); });
    // neue Lage jeder alten Zeile; Elemente und Schattierungen folgen
    std::vector<int> newIndex(item.table.size());
    for (int r = 0; r < first; ++r) newIndex[r] = r;
    std::vector<Row> sorted(item.table.begin(), item.table.begin() + first);
    for (int i = 0; i < count; ++i) {
        newIndex[rows[i]] = first + i;
        sorted.push_back(item.table[rows[i]]);
    }
    item.table = std::move(sorted);
    for (auto& it : options.gantt.items)
        if (it.row >= first && it.row < static_cast<int>(newIndex.size())) it.row = newIndex[it.row];
    for (auto& r : options.gantt.shadedRows)
        if (r >= first && r < static_cast<int>(newIndex.size())) r = newIndex[r];
    for (auto& s : options.gantt.rowStyles)
        if (s.row >= first && s.row < static_cast<int>(newIndex.size())) s.row = newIndex[s.row];
}

ShapeStyle& ganttStyle(ChartOptions& options, const GanttTarget& target) {
    if (target.item >= 0) return options.gantt.items.at(target.item).style;
    if (target.row < kFirstActivityRow && target.row != kMilestoneRow) throw std::invalid_argument("Dieses Element hat kein eigenes Aussehen");
    auto& styles = options.gantt.rowStyles;
    auto it = std::find_if(styles.begin(), styles.end(), [&](const RowStyle& s) { return s.row == target.row; });
    if (it == styles.end()) {
        styles.push_back({target.row, {}, {}});
        it = styles.end() - 1;
    }
    return target.kind == GanttTarget::Kind::Milestone ? it->milestone : it->bar;
}

void shiftGanttRows(ChartOptions& options, int from, int delta) {
    if (delta < 0) {  // gelöschte Zeilen verlieren ihre Elemente, Schattierungen und ihr Aussehen
        const int to = from - delta;
        auto gone = [&](int r) { return r >= from && r < to; };
        auto& items = options.gantt.items;
        items.erase(std::remove_if(items.begin(), items.end(), [&](const GanttItem& it) { return gone(it.row) && it.kind != GanttItem::Kind::Shade; }),
                    items.end());
        auto& shaded = options.gantt.shadedRows;
        shaded.erase(std::remove_if(shaded.begin(), shaded.end(), gone), shaded.end());
        auto& styles = options.gantt.rowStyles;
        styles.erase(std::remove_if(styles.begin(), styles.end(), [&](const RowStyle& s) { return gone(s.row); }), styles.end());
        shiftRows(options, to, delta);
        return;
    }
    shiftRows(options, from, delta);
}

void deleteGanttTarget(DataItem& item, ChartOptions& options, const GanttTarget& target) {
    if (target.kind == GanttTarget::Kind::None) return;
    if (target.item >= 0) {
        options.gantt.items.erase(options.gantt.items.begin() + target.item);
        return;
    }
    if (target.kind == GanttTarget::Kind::Milestone) {
        clearCell(item, target.row, kMilestoneCol);
    } else {
        clearCell(item, target.row, kStartCol);
        clearCell(item, target.row, kEndCol);
    }
}

}  // namespace rethink::engine
