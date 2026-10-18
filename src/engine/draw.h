#pragma once

// Hilfen zum Erzeugen von Grundformen (nur innerhalb des Layout-Kerns).

#include "chart.h"

#include <algorithm>
#include <format>
#include <string>

namespace rethink::engine::draw {

// Typografische Näherungen; die COM-Ausgabe kann später echt messen.
inline Emu textHeight(double pt) { return static_cast<Emu>(pt * 1.35 * kEmuPerPt); }
inline Emu textWidth(const std::wstring& s, double pt) { return static_cast<Emu>(s.size() * pt * 0.55 * kEmuPerPt); }

inline Color seriesColor(int row) { return Color::schemeColor(std::format(L"accent{}", (row - 1) % 6 + 1)); }

// Lesbare Schriftfarbe auf einer Füllung.
inline Color contrastOn(const Color& fill) {
    if (fill.kind != Color::Kind::Rgb) return Color::schemeColor(L"bg1");
    const double r = ((fill.rgb >> 16) & 0xFF) / 255.0, g = ((fill.rgb >> 8) & 0xFF) / 255.0, b = (fill.rgb & 0xFF) / 255.0;
    const double luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b;
    return luminance > 0.55 ? Color::schemeColor(L"tx1") : Color::schemeColor(L"bg1");
}

inline Prim text(Emu x, Emu y, Emu w, Emu h, std::wstring s, double pt, Align align, Color color, std::string role) {
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

inline Prim line(Emu x1, Emu y1, Emu x2, Emu y2, Color color, double pt, std::string role) {
    Prim p;
    p.kind = Prim::Kind::Line;
    p.x = x1, p.y = y1, p.w = x2 - x1, p.h = y2 - y1;
    p.stroke = std::move(color);
    p.strokePt = pt;
    p.role = std::move(role);
    return p;
}

inline Prim rect(Emu x, Emu y, Emu w, Emu h, Color fill, std::string role) {
    Prim p;
    p.kind = Prim::Kind::Rect;
    p.x = x, p.y = y, p.w = std::max<Emu>(w, 1), p.h = std::max<Emu>(h, 1);
    p.fill = std::move(fill);
    p.role = std::move(role);
    return p;
}

inline Prim shape(const char* preset, Emu x, Emu y, Emu w, Emu h, Color fill, std::string role, double rotation = 0) {
    Prim p = rect(x, y, w, h, std::move(fill), std::move(role));
    p.kind = Prim::Kind::Shape;
    p.shape = preset;
    p.rotation = rotation;
    return p;
}

}  // namespace rethink::engine::draw
