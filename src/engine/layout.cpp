#include "chart.h"

#include <algorithm>
#include <cmath>
#include <format>
#include <stdexcept>

namespace rethink::engine {
namespace {

// Typografische Näherungen; die COM-Ausgabe kann später echt messen.
Emu textHeight(double pt) { return static_cast<Emu>(pt * 1.35 * kEmuPerPt); }
Emu textWidth(const std::wstring& s, double pt) { return static_cast<Emu>(s.size() * pt * 0.55 * kEmuPerPt); }

Color seriesColor(int row) { return Color::schemeColor(std::format(L"accent{}", (row - 1) % 6 + 1)); }

// Lesbare Schriftfarbe auf einer Füllung.
Color contrastOn(const Color& fill) {
    if (fill.kind != Color::Kind::Rgb) return Color::schemeColor(L"bg1");
    const double r = ((fill.rgb >> 16) & 0xFF) / 255.0, g = ((fill.rgb >> 8) & 0xFF) / 255.0, b = (fill.rgb & 0xFF) / 255.0;
    const double luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b;
    return luminance > 0.55 ? Color::schemeColor(L"tx1") : Color::schemeColor(L"bg1");
}

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

Prim text(Emu x, Emu y, Emu w, Emu h, std::wstring s, double pt, Align align, Color color, std::string role) {
    Prim p;
    p.kind = Prim::Kind::Text;
    p.x = x, p.y = y, p.w = w, p.h = h;
    p.text = std::move(s);
    p.fontPt = pt;
    p.align = align;
    p.textColor = std::move(color);
    p.role = std::move(role);
    return p;
}

Prim line(Emu x1, Emu y1, Emu x2, Emu y2, Color color, double pt, std::string role) {
    Prim p;
    p.kind = Prim::Kind::Line;
    p.x = x1, p.y = y1, p.w = x2 - x1, p.h = y2 - y1;
    p.stroke = std::move(color);
    p.strokePt = pt;
    p.role = std::move(role);
    return p;
}

Drawing layoutWaterfall(const DataItem& item, const ChartOptions& o, Box b) {
    Drawing d;
    d.bounds = b;
    const auto columns = computeWaterfall(item, o, &d.warnings);
    const size_t n = columns.size();

    const int decimals = chartDecimals(item, o);
    auto fmt = [&](double v, bool percent) { return formatValue(v, decimals, o.decimalSeparator, percent); };

    // Zeilen der Datentabelle: alle nicht leeren Datenzeilen in Blattreihenfolge
    std::vector<int> tableRows;
    if (o.dataTable)
        for (size_t r = 1; r < item.table.size(); ++r)
            if (!item.table[r].empty()) tableRows.push_back(static_cast<int>(r));

    const Emu th = textHeight(o.fontPt);
    Emu namesW = 0;
    for (int r : tableRows) namesW = std::max(namesW, textWidth(cellText(item.table[r].front()), o.fontPt));
    if (namesW > 0) namesW += th / 2;

    const Emu top = o.segmentLabels ? th + th / 4 : th / 4;
    const Emu catH = o.categoryLabels ? th + th / 2 : 0;
    const Emu rowH = th + th / 3;
    const Emu tableH = static_cast<Emu>(tableRows.size()) * rowH;
    Box plot{b.x + namesW, b.y + top, b.w - namesW, b.h - top - catH - tableH};
    if (plot.w <= 0 || plot.h <= th) throw std::invalid_argument("Platzhalter ist zu klein für Diagramm und Tabelle");

    double vmin = 0, vmax = 0;
    for (auto& c : columns)
        for (auto& s : c.segments) {
            vmin = std::min(vmin, s.lo);
            vmax = std::max(vmax, s.hi);
        }
    if (vmax - vmin < 1e-12) vmax = vmin + 1;
    // Platz für Beschriftungen außerhalb negativer Säulen
    if (vmin < 0 && o.segmentLabels) vmin -= (vmax - vmin) * th / static_cast<double>(plot.h);
    auto yOf = [&](double v) { return plot.y + static_cast<Emu>(std::llround((vmax - v) / (vmax - vmin) * plot.h)); };

    const double slot = static_cast<double>(plot.w) / n;
    const Emu colW = static_cast<Emu>(slot / (1.0 + o.gap));
    auto colX = [&](size_t j) { return plot.x + static_cast<Emu>(j * slot + (slot - colW) / 2.0); };

    // Segmente und Beschriftungen
    for (size_t j = 0; j < n; ++j) {
        for (auto& s : columns[j].segments) {
            Prim rect;
            rect.kind = Prim::Kind::Rect;
            rect.x = colX(j);
            rect.y = yOf(s.hi);
            rect.w = colW;
            rect.h = std::max<Emu>(yOf(s.lo) - yOf(s.hi), 1);
            rect.fill = s.fill ? Color::rgbColor(*s.fill) : seriesColor(s.row);
            rect.role = std::format("segment r{} c{}", s.row, j + 1);
            rect.flags = s.computed ? "e" : s.value < 0 ? "neg" : "pos";
            rect.value = s.value;
            d.prims.push_back(rect);

            if (!o.segmentLabels) continue;
            const std::wstring label = fmt(s.value, s.percent);
            const bool fits = rect.h >= th && colW >= textWidth(label, o.fontPt) + th / 4;
            if (fits) {
                d.prims.push_back(text(rect.x, rect.y, rect.w, rect.h, label, o.fontPt, Align::Center, contrastOn(rect.fill),
                                       std::format("label r{} c{}", s.row, j + 1)));
            } else {
                const bool above = s.value >= 0;
                const Emu y = above ? rect.y - th : rect.y + rect.h;
                d.prims.push_back(text(rect.x - colW / 2, y, colW * 2, th, label, o.fontPt, Align::Center, Color::schemeColor(L"tx1"),
                                       std::format("label r{} c{}", s.row, j + 1)));
            }
        }
    }

    // Verbinder zwischen benachbarten Säulen
    if (o.connectors) {
        for (size_t j = 0; j + 1 < n; ++j) {
            const Emu y = yOf(columns[j].levelAfter);
            Prim c = line(colX(j) + colW, y, colX(j + 1), y, Color::schemeColor(L"tx1"), 0.5, std::format("connector c{}", j + 1));
            d.prims.push_back(c);
        }
    }

    // Basislinie
    d.prims.push_back(line(plot.x, yOf(0), plot.x + plot.w, yOf(0), Color::schemeColor(L"tx1"), 0.75, "baseline"));

    // Kategoriebeschriftungen
    const Emu plotBottom = plot.y + plot.h;
    if (o.categoryLabels) {
        const Row& head = item.table.front();
        for (size_t j = 0; j < n; ++j) {
            const std::wstring label = j + 1 < head.size() ? cellText(head[j + 1]) : L"";
            d.prims.push_back(text(plot.x + static_cast<Emu>(j * slot), plotBottom + th / 4, static_cast<Emu>(slot), th, label, o.fontPt,
                                   Align::Center, Color::schemeColor(L"tx1"), std::format("category c{}", j + 1)));
        }
    }

    // Datentabelle: an den Säulen ausgerichtet, "e" zeigt den berechneten Wert
    Emu y = plotBottom + catH;
    for (int r : tableRows) {
        const Row& row = item.table[r];
        d.prims.push_back(line(b.x, y, plot.x + plot.w, y, Color::rgbColor(0xBFBFBF), 0.5, std::format("table rule r{}", r)));
        d.prims.push_back(text(b.x, y, namesW, rowH, cellText(row.front()), o.fontPt, Align::Left, Color::schemeColor(L"tx1"),
                               std::format("table name r{}", r)));
        for (size_t j = 0; j < n; ++j) {
            std::wstring value;
            if (j + 1 < row.size()) {
                const Cell& cell = row[j + 1];
                if (isComputedCell(cell)) {
                    for (auto& s : columns[j].segments)
                        if (s.row == r && s.computed) value = fmt(s.value, s.percent);
                } else if (cell.kind == CellKind::Number || cell.kind == CellKind::Percentage) {
                    value = fmt(cell.number, cell.kind == CellKind::Percentage);
                } else {
                    value = cellText(cell);
                }
            }
            d.prims.push_back(text(plot.x + static_cast<Emu>(j * slot), y, static_cast<Emu>(slot), rowH, value, o.fontPt, Align::Center,
                                   Color::schemeColor(L"tx1"), std::format("table cell r{} c{}", r, j + 1)));
        }
        y += rowH;
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
    auto sameColor = [](const Color& x, const Color& y) { return x.kind == y.kind && x.scheme == y.scheme && x.rgb == y.rgb; };
    return a.kind == b.kind && a.x == b.x && a.y == b.y && a.w == b.w && a.h == b.h && a.text == b.text && a.role == b.role &&
           a.flags == b.flags && a.value == b.value && sameColor(a.fill, b.fill) && sameColor(a.stroke, b.stroke) &&
           sameColor(a.textColor, b.textColor) && a.fontPt == b.fontPt && a.strokePt == b.strokePt;
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
    double best = 0, scale = 0;
    for (auto& p : drawing.prims) {
        if (p.kind != Prim::Kind::Rect || p.flags.empty() || std::fabs(p.value) <= best) continue;
        best = std::fabs(p.value);
        scale = static_cast<double>(p.h) / best;
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
        case ChartType::Waterfall: return layoutWaterfall(item, options, bounds);
    }
    throw std::invalid_argument("unbekannter Diagrammtyp");
}

}  // namespace rethink::engine
