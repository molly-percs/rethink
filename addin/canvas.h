#pragma once

#include <windows.h>

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <string>

namespace rethink::addin {

// Zeichenfläche für Ebenenfenster (UpdateLayeredWindow): vorbelegte BGRA-Pixel, weiche Kanten von Hand.
// Gemeinsam für das Auswahl-Overlay und das Rundmenü des Zeitplans.

constexpr COLORREF kBlue = RGB(37, 120, 235);  // rethink-Blau

// Vorbelegtes (premultiplied) BGRA-Pixel
inline std::uint32_t pixel(COLORREF color, double alpha) {
    const auto a = static_cast<std::uint32_t>(std::clamp(alpha, 0.0, 1.0) * 255.0 + 0.5);
    const auto r = GetRValue(color) * a / 255, g = GetGValue(color) * a / 255, b = GetBValue(color) * a / 255;
    return (a << 24) | (r << 16) | (g << 8) | b;
}

class Canvas {
public:
    Canvas(std::uint32_t* px, int w, int h) : px_(px), w_(w), h_(h) {}
    void blend(int x, int y, COLORREF c, double alpha) {
        if (x < 0 || y < 0 || x >= w_ || y >= h_ || alpha <= 0) return;
        std::uint32_t& d = px_[y * w_ + x];
        const double da = (d >> 24) / 255.0, a = std::min(alpha, 1.0);
        const double out = a + da * (1 - a);
        auto channel = [&](int shift, int src) {
            const double dc = ((d >> shift) & 0xFF) / 255.0;  // vorbelegt
            return static_cast<std::uint32_t>(std::clamp((src / 255.0) * a + dc * (1 - a), 0.0, 1.0) * 255 + 0.5);
        };
        d = (static_cast<std::uint32_t>(out * 255 + 0.5) << 24) | (channel(16, GetRValue(c)) << 16) | (channel(8, GetGValue(c)) << 8) |
            channel(0, GetBValue(c));
    }
    void fill(const RECT& r, COLORREF c, double alpha) {
        for (int y = r.top; y < r.bottom; ++y)
            for (int x = r.left; x < r.right; ++x) blend(x, y, c, alpha);
    }
    // Dreieck mit Spitze nach oben (0), unten (1), links (2) oder rechts (3), weiße Kontur für Sichtbarkeit auf Segmenten
    void triangle(const RECT& r, int pointing) {
        const bool vertical = pointing <= 1;
        const int along0 = vertical ? r.top : r.left, along1 = vertical ? r.bottom : r.right;
        const double center = vertical ? (r.left + r.right) / 2.0 : (r.top + r.bottom) / 2.0;
        const double half = vertical ? (r.right - r.left) / 2.0 : (r.bottom - r.top) / 2.0;
        for (int a = along0; a < along1; ++a) {
            const double t = (a + 0.5 - along0) / (along1 - along0);  // 0 … 1 entlang der Spitze-Basis-Achse
            const double w = half * (pointing == 0 || pointing == 2 ? t : 1 - t);
            const int c0 = vertical ? r.left : r.top, c1 = vertical ? r.right : r.bottom;
            for (int c = c0; c < c1; ++c) {
                const double d = std::fabs(c + 0.5 - center);
                const int x = vertical ? c : a, y = vertical ? a : c;
                if (d <= w - 1.0) blend(x, y, kBlue, 1.0);
                else if (d <= w + 0.6) blend(x, y, RGB(255, 255, 255), 0.95);
            }
        }
    }
    // Schmaler Griff mit weißem Rand (Balkenende)
    void grip(const RECT& r) {
        fill(r, RGB(255, 255, 255), 0.95);
        fill({r.left + 1, r.top + 1, r.right - 1, r.bottom - 1}, kBlue, 1.0);
    }
    int width() const { return w_; }
    int height() const { return h_; }

private:
    std::uint32_t* px_;
    int w_, h_;
};

// GDI-Text in eine eigene Maske zeichnen und als Alpha in Farbe übernehmen (die Ebene hat keinen GDI-Alphakanal).
inline void drawText(Canvas& canvas, HDC screen, HFONT font, const std::wstring& text, RECT rect, UINT format, COLORREF color, double alpha) {
    BITMAPINFO info{};
    info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    info.bmiHeader.biWidth = canvas.width();
    info.bmiHeader.biHeight = -canvas.height();
    info.bmiHeader.biPlanes = 1;
    info.bmiHeader.biBitCount = 32;
    info.bmiHeader.biCompression = BI_RGB;
    HDC maskDc = CreateCompatibleDC(screen);
    void* maskBits = nullptr;
    HBITMAP mask = CreateDIBSection(maskDc, &info, DIB_RGB_COLORS, &maskBits, nullptr, 0);
    if (mask) {
        HGDIOBJ oldMask = SelectObject(maskDc, mask);
        HGDIOBJ oldFont = SelectObject(maskDc, font);
        SetBkMode(maskDc, TRANSPARENT);
        SetTextColor(maskDc, RGB(255, 255, 255));
        DrawTextW(maskDc, text.c_str(), -1, &rect, format | DT_SINGLELINE | DT_NOPREFIX);
        const auto* mp = static_cast<const std::uint32_t*>(maskBits);
        const int w = canvas.width();
        for (int y = std::max<int>(rect.top, 0); y < std::min<int>(rect.bottom, canvas.height()); ++y)
            for (int x = std::max<int>(rect.left, 0); x < std::min<int>(rect.right, w); ++x) {
                const double coverage = (mp[y * w + x] & 0xFF) / 255.0;
                if (coverage > 0) canvas.blend(x, y, color, alpha * coverage);
            }
        SelectObject(maskDc, oldFont);
        SelectObject(maskDc, oldMask);
        DeleteObject(mask);
    }
    DeleteDC(maskDc);
}

}  // namespace rethink::addin
