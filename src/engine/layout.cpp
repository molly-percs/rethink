#include "chart.h"
#include "draw.h"

#include <algorithm>
#include <cmath>
#include <format>
#include <stdexcept>

namespace rethink::engine {

Drawing layoutGantt(const DataItem& item, const ChartOptions& options, Box bounds);  // gantt.cpp

namespace {

using namespace draw;

int autoDecimals(const std::vector<double>& values) {
    int decimals = 0;
    for (double v : values) {
        for (int d = 0; d <= 2; ++d) {
            const double scaled = v * std::pow(10.0, d);
            if (std::fabs(scaled - std::round(scaled)) < 1e-9) {
                decimals = std::max(decimals, d);
                break;
            }
            if (d == 2) decimals = 2;
        }
    }
    return decimals;
}

// Gemeinsame Teile beider Wasserfall-Ausrichtungen
struct WaterfallBase {
    std::vector<WaterfallColumn> columns;
    std::vector<int> tableRows;  // Zeilen der Datentabelle: alle nicht leeren Datenzeilen in Blattreihenfolge
    int decimals = 0;
    double vmin = 0, vmax = 0;
};

WaterfallBase waterfallBase(const DataItem& item, const ChartOptions& o, Drawing& d) {
    WaterfallBase w;
    w.columns = computeWaterfall(item, o, &d.warnings);
    w.decimals = chartDecimals(item, o);
    if (o.dataTable)
        for (size_t r = 1; r < item.table.size(); ++r)
            if (!item.table[r].empty()) w.tableRows.push_back(static_cast<int>(r));
    for (auto& c : w.columns)
        for (auto& s : c.segments) {
            w.vmin = std::min(w.vmin, s.lo);
            w.vmax = std::max(w.vmax, s.hi);
        }
    if (w.vmax - w.vmin < 1e-12) w.vmax = w.vmin + 1;
    return w;
}

// Text einer Tabellenzelle: "e" zeigt den berechneten Wert
std::wstring tableValue(const DataItem& item, const WaterfallBase& w, const ChartOptions& o, int r, size_t j) {
    const Row& row = item.table[r];
    if (j + 1 >= row.size()) return {};
    const Cell& cell = row[j + 1];
    if (isComputedCell(cell)) {
        for (auto& s : w.columns[j].segments)
            if (s.row == r && s.computed) return formatValue(s.value, w.decimals, o.decimalSeparator, s.percent);
        return {};
    }
    if (cell.kind == CellKind::Number || cell.kind == CellKind::Percentage)
        return formatValue(cell.number, w.decimals, o.decimalSeparator, cell.kind == CellKind::Percentage);
    return cellText(cell);
}

Prim segmentRect(const WaterfallSegment& s, size_t j, Emu x, Emu y, Emu w, Emu h) {
    Prim p = rect(x, y, w, h, s.fill ? Color::rgbColor(*s.fill) : seriesColor(s.row), std::format("segment r{} c{}", s.row, j + 1));
    p.flags = s.computed ? "e" : s.value < 0 ? "neg" : "pos";
    p.value = s.value;
    return p;
}

std::wstring categoryText(const DataItem& item, size_t j) {
    const Row& head = item.table.front();
    return j + 1 < head.size() ? cellText(head[j + 1]) : L"";
}

// Säulen nach oben (Up) oder an der Waagerechten gespiegelt (Down: Werte wachsen nach unten, Kategorien oben).
Drawing layoutWaterfallVertical(const DataItem& item, const ChartOptions& o, Box b) {
    Drawing d;
    d.bounds = b;
    const bool down = o.orientation == Orientation::Down;
    d.valueDirection = o.orientation;
    WaterfallBase w = waterfallBase(item, o, d);
    const size_t n = w.columns.size();

    const Emu th = textHeight(o.fontPt);
    Emu namesW = 0;
    for (int r : w.tableRows) namesW = std::max(namesW, textWidth(cellText(item.table[r].front()), o.fontPt));
    if (namesW > 0) namesW += th / 2;

    const Emu labelPad = o.segmentLabels ? th + th / 4 : th / 4;  // Platz für Beschriftungen außerhalb (Werteseite)
    const Emu catH = o.categoryLabels ? th + th / 2 : 0;
    const Emu rowH = th + th / 3;
    const Emu tableH = static_cast<Emu>(w.tableRows.size()) * rowH;
    const Emu plotH = b.h - labelPad - catH - tableH;
    Box plot{b.x + namesW, down ? b.y + catH : b.y + labelPad, b.w - namesW, plotH};
    if (plot.w <= 0 || plot.h <= th) throw std::invalid_argument("Platzhalter ist zu klein für Diagramm und Tabelle");

    // Platz für Beschriftungen außerhalb negativer Säulen
    if (w.vmin < 0 && o.segmentLabels) w.vmin -= (w.vmax - w.vmin) * th / static_cast<double>(plot.h);
    auto yOf = [&](double v) {
        const double f = down ? (v - w.vmin) / (w.vmax - w.vmin) : (w.vmax - v) / (w.vmax - w.vmin);
        return plot.y + static_cast<Emu>(std::llround(f * plot.h));
    };

    const double slot = static_cast<double>(plot.w) / n;
    const Emu colW = static_cast<Emu>(slot / (1.0 + o.gap));
    auto colX = [&](size_t j) { return plot.x + static_cast<Emu>(j * slot + (slot - colW) / 2.0); };

    for (size_t j = 0; j < n; ++j) {
        for (auto& s : w.columns[j].segments) {
            const Emu y0 = yOf(s.lo), y1 = yOf(s.hi);
            const Prim r = segmentRect(s, j, colX(j), std::min(y0, y1), colW, std::abs(y0 - y1));
            d.prims.push_back(r);
            if (!o.segmentLabels) continue;
            const std::wstring label = formatValue(s.value, w.decimals, o.decimalSeparator, s.percent);
            const std::string role = std::format("label r{} c{}", s.row, j + 1);
            if (r.h >= th && colW >= textWidth(label, o.fontPt) + th / 4) {
                d.prims.push_back(text(r.x, r.y, r.w, r.h, label, o.fontPt, Align::Center, contrastOn(r.fill), role));
            } else {
                const bool towardsTop = (s.value >= 0) != down;  // außerhalb am wachsenden Ende
                const Emu y = towardsTop ? r.y - th : r.y + r.h;
                d.prims.push_back(text(r.x - colW / 2, y, colW * 2, th, label, o.fontPt, Align::Center, Color::schemeColor(L"tx1"), role));
            }
        }
    }

    if (o.connectors)
        for (size_t j = 0; j + 1 < n; ++j) {
            const Emu y = yOf(w.columns[j].levelAfter);
            d.prims.push_back(line(colX(j) + colW, y, colX(j + 1), y, Color::schemeColor(L"tx1"), 0.5, std::format("connector c{}", j + 1)));
        }

    d.prims.push_back(line(plot.x, yOf(0), plot.x + plot.w, yOf(0), Color::schemeColor(L"tx1"), 0.75, "baseline"));

    const Emu plotBottom = plot.y + plot.h;
    if (o.categoryLabels) {
        const Emu y = down ? b.y + th / 4 : plotBottom + th / 4;
        for (size_t j = 0; j < n; ++j)
            d.prims.push_back(text(plot.x + static_cast<Emu>(j * slot), y, static_cast<Emu>(slot), th, categoryText(item, j), o.fontPt, Align::Center,
                                   Color::schemeColor(L"tx1"), std::format("category c{}", j + 1)));
    }

    // Datentabelle unter dem Diagramm, an den Säulen ausgerichtet
    Emu y = down ? plotBottom + labelPad : plotBottom + catH;
    for (int r : w.tableRows) {
        const Row& row = item.table[r];
        d.prims.push_back(line(b.x, y, plot.x + plot.w, y, Color::rgbColor(0xBFBFBF), 0.5, std::format("table rule r{}", r)));
        d.prims.push_back(text(b.x, y, namesW, rowH, cellText(row.front()), o.fontPt, Align::Left, Color::schemeColor(L"tx1"),
                               std::format("table name r{}", r)));
        for (size_t j = 0; j < n; ++j)
            d.prims.push_back(text(plot.x + static_cast<Emu>(j * slot), y, static_cast<Emu>(slot), rowH, tableValue(item, w, o, r, j), o.fontPt,
                                   Align::Center, Color::schemeColor(L"tx1"), std::format("table cell r{} c{}", r, j + 1)));
        y += rowH;
    }
    return d;
}

// Balken nach rechts (Right) oder links (Left): Kategorien von oben nach unten, Beschriftungen auf der Seite der
// Basislinie; die Datentabelle steht als Spalten auf der anderen Seite, Zeilennamen als Überschrift.
Drawing layoutWaterfallHorizontal(const DataItem& item, const ChartOptions& o, Box b) {
    Drawing d;
    d.bounds = b;
    const bool left = o.orientation == Orientation::Left;
    d.valueDirection = o.orientation;
    WaterfallBase w = waterfallBase(item, o, d);
    const size_t n = w.columns.size();

    const Emu th = textHeight(o.fontPt);
    Emu catW = 0;
    if (o.categoryLabels) {
        for (size_t j = 0; j < n; ++j) catW = std::max(catW, textWidth(categoryText(item, j), o.fontPt));
        catW += th / 2;
    }
    // Tabellenspalten: so breit wie der längste Eintrag
    std::vector<Emu> tableW;
    Emu tableTotal = 0;
    for (int r : w.tableRows) {
        Emu width = textWidth(cellText(item.table[r].front()), o.fontPt);
        for (size_t j = 0; j < n; ++j) width = std::max(width, textWidth(tableValue(item, w, o, r, j), o.fontPt));
        tableW.push_back(width + th / 2);
        tableTotal += tableW.back();
    }
    const Emu headH = w.tableRows.empty() ? 0 : th + th / 3;
    Emu labelW = 0;  // Platz für Beschriftungen außerhalb
    if (o.segmentLabels) {
        for (auto& c : w.columns)
            for (auto& s : c.segments) labelW = std::max(labelW, textWidth(formatValue(s.value, w.decimals, o.decimalSeparator, s.percent), o.fontPt));
        labelW += th / 4;
    }
    const Emu growPad = labelW, shrinkPad = w.vmin < 0 ? labelW : th / 4;
    const Emu plotW = b.w - catW - tableTotal - growPad - shrinkPad;
    Box plot{left ? b.x + tableTotal + growPad : b.x + catW + shrinkPad, b.y + headH, plotW, b.h - headH};
    if (plot.w <= th || plot.h <= th) throw std::invalid_argument("Platzhalter ist zu klein für Diagramm und Tabelle");

    auto xOf = [&](double v) {
        const double f = left ? (w.vmax - v) / (w.vmax - w.vmin) : (v - w.vmin) / (w.vmax - w.vmin);
        return plot.x + static_cast<Emu>(std::llround(f * plot.w));
    };
    const double slot = static_cast<double>(plot.h) / n;
    const Emu barH = static_cast<Emu>(slot / (1.0 + o.gap));
    auto barY = [&](size_t j) { return plot.y + static_cast<Emu>(j * slot + (slot - barH) / 2.0); };

    for (size_t j = 0; j < n; ++j) {
        for (auto& s : w.columns[j].segments) {
            const Emu x0 = xOf(s.lo), x1 = xOf(s.hi);
            const Prim r = segmentRect(s, j, std::min(x0, x1), barY(j), std::abs(x1 - x0), barH);
            d.prims.push_back(r);
            if (!o.segmentLabels) continue;
            const std::wstring label = formatValue(s.value, w.decimals, o.decimalSeparator, s.percent);
            const std::string role = std::format("label r{} c{}", s.row, j + 1);
            const Emu lw = textWidth(label, o.fontPt) + th / 4;
            if (barH >= th && r.w >= lw) {
                d.prims.push_back(text(r.x, r.y, r.w, r.h, label, o.fontPt, Align::Center, contrastOn(r.fill), role));
            } else {
                const bool towardsRight = (s.value >= 0) != left;
                const Emu x = towardsRight ? r.x + r.w + th / 8 : r.x - lw;
                d.prims.push_back(text(x, r.y + (barH - th) / 2, lw, th, label, o.fontPt, towardsRight ? Align::Left : Align::Right,
                                       Color::schemeColor(L"tx1"), role));
            }
        }
    }

    if (o.connectors)
        for (size_t j = 0; j + 1 < n; ++j) {
            const Emu x = xOf(w.columns[j].levelAfter);
            d.prims.push_back(line(x, barY(j) + barH, x, barY(j + 1), Color::schemeColor(L"tx1"), 0.5, std::format("connector c{}", j + 1)));
        }

    d.prims.push_back(line(xOf(0), plot.y, xOf(0), plot.y + plot.h, Color::schemeColor(L"tx1"), 0.75, "baseline"));

    if (o.categoryLabels) {
        const Emu x = left ? b.x + b.w - catW + th / 4 : b.x;
        for (size_t j = 0; j < n; ++j)
            d.prims.push_back(text(x, plot.y + static_cast<Emu>(j * slot), catW - th / 4, static_cast<Emu>(slot), categoryText(item, j), o.fontPt,
                                   left ? Align::Left : Align::Right, Color::schemeColor(L"tx1"), std::format("category c{}", j + 1)));
    }

    // Datentabelle als Spalten neben dem Diagramm, an den Balken ausgerichtet
    Emu x = left ? b.x : b.x + b.w - tableTotal;
    for (size_t t = 0; t < w.tableRows.size(); ++t) {
        const int r = w.tableRows[t];
        d.prims.push_back(line(x, b.y, x, plot.y + plot.h, Color::rgbColor(0xBFBFBF), 0.5, std::format("table rule r{}", r)));
        d.prims.push_back(text(x, b.y, tableW[t], headH, cellText(item.table[r].front()), o.fontPt, Align::Center, Color::schemeColor(L"tx1"),
                               std::format("table name r{}", r)));
        for (size_t j = 0; j < n; ++j)
            d.prims.push_back(text(x, plot.y + static_cast<Emu>(j * slot), tableW[t], static_cast<Emu>(slot), tableValue(item, w, o, r, j), o.fontPt,
                                   Align::Center, Color::schemeColor(L"tx1"), std::format("table cell r{} c{}", r, j + 1)));
        x += tableW[t];
    }
    return d;
}

}  // namespace

std::wstring formatValue(double value, int decimals, const std::wstring& decimalSeparator, bool percent) {
    std::wstring s = std::format(L"{:.{}f}", value, decimals);
    if (s == L"-0" || s.starts_with(L"-0.") && std::all_of(s.begin() + 3, s.end(), [](wchar_t c) { return c == L'0'; })) s.erase(0, 1);
    if (const auto dot = s.find(L'.'); dot != std::wstring::npos) s.replace(dot, 1, decimalSeparator);
    return percent ? s + L"%" : s;
}

bool samePrim(const Prim& a, const Prim& b) {
    auto sameColor = [](const Color& x, const Color& y) {
        return x.kind == y.kind && x.scheme == y.scheme && x.rgb == y.rgb && x.brightness == y.brightness;
    };
    return a.kind == b.kind && a.x == b.x && a.y == b.y && a.w == b.w && a.h == b.h && a.text == b.text && a.role == b.role &&
           a.flags == b.flags && a.value == b.value && sameColor(a.fill, b.fill) && sameColor(a.stroke, b.stroke) &&
           sameColor(a.textColor, b.textColor) && a.fontPt == b.fontPt && a.strokePt == b.strokePt && a.bold == b.bold && a.shape == b.shape &&
           a.rotation == b.rotation && a.dashed == b.dashed && a.align == b.align && a.transparency == b.transparency;
}

int stepDecimals(const Cell& cell, const ChartOptions& options) {
    if (options.decimals) return *options.decimals;
    const std::string s = std::format("{}", cell.number);
    if (s.find_first_of("eE") != std::string::npos) return 0;
    const auto dot = s.find('.');
    return dot == std::string::npos ? 0 : std::min(4, static_cast<int>(s.size() - dot - 1));
}

double dragValue(double startValue, double pixels, double pixelsPerUnit, int decimals) {
    const double scale = std::pow(10.0, decimals);
    const double raw = startValue + (pixelsPerUnit > 0 ? pixels / pixelsPerUnit : 0);
    const double rounded = std::round(raw * scale) / scale;
    return rounded == 0.0 ? 0.0 : rounded;  // kein -0
}

double emuPerUnit(const Drawing& drawing) {
    const bool horizontal = drawing.valueDirection == Orientation::Right || drawing.valueDirection == Orientation::Left;
    double best = 0, scale = 0;
    for (auto& p : drawing.prims) {
        if (p.kind != Prim::Kind::Rect || p.flags.empty() || std::fabs(p.value) <= best) continue;
        best = std::fabs(p.value);
        scale = static_cast<double>(horizontal ? p.w : p.h) / best;
    }
    return scale;
}

int chartDecimals(const DataItem& item, const ChartOptions& options) {
    if (options.decimals) return *options.decimals;
    std::vector<double> values;
    for (size_t r = 1; r < item.table.size(); ++r)
        for (size_t c = 1; c < item.table[r].size(); ++c) {
            const Cell& cell = item.table[r][c];
            if (cell.kind == CellKind::Number || cell.kind == CellKind::Percentage) values.push_back(cell.number);
        }
    return autoDecimals(values);
}

Drawing layoutChart(const DataItem& item, const ChartOptions& options, Box bounds) {
    switch (options.type) {
        case ChartType::Waterfall: {
            Drawing d = options.orientation == Orientation::Right || options.orientation == Orientation::Left
                            ? layoutWaterfallHorizontal(item, options, bounds)
                            : layoutWaterfallVertical(item, options, bounds);
            // Unsichtbare Fläche: ein Klick irgendwo ins Diagramm trifft das Element (sonst nur seine Formen)
            Prim background = rect(bounds.x, bounds.y, bounds.w, bounds.h, Color::schemeColor(L"bg1"), "background");
            background.transparency = 1;
            d.prims.insert(d.prims.begin(), background);
            return d;
        }
        case ChartType::Gantt: return layoutGantt(item, options, bounds);
    }
    throw std::invalid_argument("unbekannter Diagrammtyp");
}

}  // namespace rethink::engine
