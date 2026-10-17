#include "chart.h"

#include <algorithm>
#include <stdexcept>

namespace rethink::engine {
namespace {

bool isTableOnly(const ChartOptions& options, int row) {
    return std::find(options.tableRows.begin(), options.tableRows.end(), row) != options.tableRows.end();
}

// Stapelt Segmente ab `base`: positive nach oben, negative nach unten.
void stack(std::vector<WaterfallSegment>& segments, double base) {
    double up = base, down = base;
    for (auto& s : segments) {
        if (s.value >= 0) {
            s.lo = up;
            s.hi = up + s.value;
            up = s.hi;
        } else {
            s.hi = down;
            s.lo = down + s.value;
            down = s.lo;
        }
    }
}

double sumExcept(const std::vector<WaterfallSegment>& segments, bool computed) {
    double sum = 0;
    for (auto& s : segments)
        if (s.computed == computed) sum += s.value;
    return sum;
}

}  // namespace

bool isComputedCell(const Cell& cell) {
    return cell.kind == CellKind::String && (cell.text == L"e" || cell.text == L"E");
}

std::vector<WaterfallColumn> computeWaterfall(const DataItem& item, const ChartOptions& options,
                                              std::vector<std::string>* warnings) {
    if (item.table.empty()) throw std::invalid_argument("\"table\" ist leer");
    size_t columns = item.table.front().size() > 0 ? item.table.front().size() - 1 : 0;
    for (size_t r = 1; r < item.table.size(); ++r)
        if (!isTableOnly(options, static_cast<int>(r)) && item.table[r].size() > 1) columns = std::max(columns, item.table[r].size() - 1);
    if (columns == 0) throw std::invalid_argument("keine Datenspalten");

    std::vector<WaterfallColumn> result(columns);
    bool warned = false;
    for (size_t j = 0; j < columns; ++j) {
        for (size_t r = 1; r < item.table.size(); ++r) {
            const Row& row = item.table[r];
            if (row.empty() || isTableOnly(options, static_cast<int>(r)) || j + 1 >= row.size()) continue;
            const Cell& cell = row[j + 1];
            WaterfallSegment s;
            s.row = static_cast<int>(r);
            s.fill = cell.fill;
            if (isComputedCell(cell)) {
                s.computed = true;
            } else if (cell.kind == CellKind::Number || cell.kind == CellKind::Percentage) {
                s.value = cell.number;
                s.percent = cell.kind == CellKind::Percentage;
            } else {
                if (cell.kind != CellKind::Empty && warnings && !warned) {
                    warnings->push_back("Text- und Datumswerte in Wasserfall-Datenzellen werden ignoriert (nur Zahlen und \"e\")");
                    warned = true;
                }
                continue;
            }
            result[j].segments.push_back(s);
        }
        const auto computed = std::count_if(result[j].segments.begin(), result[j].segments.end(), [](auto& s) { return s.computed; });
        if (computed > 1) throw std::invalid_argument("höchstens ein \"e\" pro Säule (Spalte " + std::to_string(j + 1) + ")");
        result[j].anchored = computed == 1;
    }

    // "e" in der ersten Säule: rückwärts rechnen, sodass der Stand am Ende 0 ist.
    double level = 0;
    if (result.front().anchored) {
        double increments = 0;
        for (size_t j = 1; j < columns; ++j)
            if (!result[j].anchored) increments += sumExcept(result[j].segments, false);
        level = -increments;  // Höhe der ersten Säule
    }

    for (size_t j = 0; j < columns; ++j) {
        WaterfallColumn& col = result[j];
        col.levelBefore = j == 0 ? 0 : level;
        if (col.anchored) {
            const double top = level;  // bei j == 0 bereits rückwärts bestimmt
            const double others = sumExcept(col.segments, false);
            for (auto& s : col.segments)
                if (s.computed) s.value = top - others;
            stack(col.segments, 0);
            level = top;
        } else {
            stack(col.segments, level);
            level += sumExcept(col.segments, false);
        }
        col.levelAfter = level;
    }
    return result;
}

}  // namespace rethink::engine
