#include "overlay.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <cstdlib>
#include <vector>

extern "C" IMAGE_DOS_HEADER __ImageBase;

namespace rethink::addin {
namespace {

constexpr wchar_t kClass[] = L"rethinkSelectionOverlay";
constexpr UINT kMsgAction = WM_APP + 41;  // wParam = Geste, lParam = Mausweg
constexpr COLORREF kBlue = RGB(37, 120, 235);  // rethink-Blau

HWND g_window = nullptr;
OverlayContent g_content;
POINT g_origin{};  // Bildschirmposition der linken oberen Fensterecke
bool g_displaced = false;
OverlayCallback g_callback = nullptr;

// Gedrückter Griff. Kopie statt Index: während des Ziehens baut der Rückruf das Overlay laufend neu auf.
struct Press {
    bool active = false;
    OverlayHandle handle;
    POINT start{};
    bool shift = false;
    bool dragging = false;
    bool posted = false;  // eine Drag-Nachricht steht noch aus: nur den neuesten Mausweg merken
    int pixels = 0;
};
Press g_press;

int scaled(HWND hwnd, int px) {
    const UINT dpi = hwnd ? GetDpiForWindow(hwnd) : 96;
    return MulDiv(px, dpi ? dpi : 96, 96);
}

// Rand des Overlay-Fensters um das Element (Leuchten, Griffe, Symbol, Etikett)
int margin(HWND hwnd) { return scaled(hwnd, 34); }

int hitHandle(POINT screen) {
    for (size_t i = 0; i < g_content.handles.size(); ++i)
        if (PtInRect(&g_content.handles[i].rect, screen)) return static_cast<int>(i);
    return -1;
}

POINT messagePos() {
    const DWORD pos = GetMessagePos();  // Bildschirmkoordinaten zum Zeitpunkt der Nachricht
    return {static_cast<short>(LOWORD(pos)), static_cast<short>(HIWORD(pos))};
}

void post(HWND hwnd, OverlayGesture gesture, int pixels) {
    PostMessageW(hwnd, kMsgAction, static_cast<WPARAM>(gesture), static_cast<LPARAM>(pixels));
}

// Drücken beenden; ein laufendes Ziehen endet mit `gesture` (DragEnd oder DragCancel).
void release(HWND hwnd, OverlayGesture gesture) {
    if (!g_press.active) return;
    const bool dragging = g_press.dragging;
    g_press.active = false;  // vor ReleaseCapture: WM_CAPTURECHANGED soll nichts mehr auslösen
    if (GetCapture() == hwnd) ReleaseCapture();
    if (dragging) post(hwnd, gesture, g_press.pixels);
    else if (gesture == OverlayGesture::DragEnd) post(hwnd, OverlayGesture::Click, 0);
}

LRESULT CALLBACK windowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_NCHITTEST: {
            // Nur Griffe und Symbol sind anklickbar; alles andere geht an PowerPoint (gleicher Thread).
            const POINT p{static_cast<short>(LOWORD(lParam)), static_cast<short>(HIWORD(lParam))};
            return hitHandle(p) >= 0 ? HTCLIENT : HTTRANSPARENT;
        }
        case WM_MOUSEACTIVATE: return MA_NOACTIVATE;
        case WM_SETCURSOR: {
            const int index = hitHandle(messagePos());
            if (index < 0 || g_content.handles[index].style == HandleStyle::Datasheet) {
                SetCursor(LoadCursorW(nullptr, IDC_HAND));
            } else {
                const OverlayHandle& h = g_content.handles[index];
                SetCursor(LoadCursorW(nullptr, h.dirX != 0 ? IDC_SIZEWE : IDC_SIZENS));
            }
            return TRUE;
        }
        case WM_LBUTTONDOWN: {
            const POINT p = messagePos();
            const int index = hitHandle(p);
            if (index < 0) return 0;
            g_press = {};
            g_press.handle = g_content.handles[index];
            g_press.shift = (wParam & MK_SHIFT) != 0;
            if (g_press.handle.action == OverlayAction::Datasheet) {
                post(hwnd, OverlayGesture::Click, 0);
                return 0;
            }
            // Klick oder Ziehen entscheidet sich erst an Bewegung bzw. Loslassen
            g_press.active = true;
            g_press.start = p;
            SetCapture(hwnd);
            return 0;
        }
        case WM_MOUSEMOVE: {
            if (!g_press.active) return 0;
            if (GetAsyncKeyState(VK_ESCAPE) & 0x8000) {
                release(hwnd, OverlayGesture::DragCancel);
                return 0;
            }
            const POINT p = messagePos();
            const OverlayHandle& h = g_press.handle;
            const int pixels = (p.x - g_press.start.x) * h.dirX + (p.y - g_press.start.y) * h.dirY;  // entlang der Griffrichtung
            if (!g_press.dragging && std::abs(pixels) < GetSystemMetrics(h.dirX != 0 ? SM_CXDRAG : SM_CYDRAG)) return 0;
            g_press.dragging = true;
            g_press.pixels = pixels;
            SetCursor(LoadCursorW(nullptr, h.dirX != 0 ? IDC_SIZEWE : IDC_SIZENS));
            if (!g_press.posted) {
                g_press.posted = true;
                post(hwnd, OverlayGesture::Drag, pixels);
            }
            return 0;
        }
        case WM_LBUTTONUP: release(hwnd, OverlayGesture::DragEnd); return 0;
        case WM_CAPTURECHANGED: release(hwnd, OverlayGesture::DragCancel); return 0;
        case kMsgAction: {
            const auto gesture = static_cast<OverlayGesture>(wParam);
            int pixels = static_cast<int>(lParam);
            if (gesture == OverlayGesture::Drag) {
                if (!g_press.posted) return 0;  // schon durch DragEnd/DragCancel überholt
                g_press.posted = false;
                pixels = g_press.pixels;  // neuester Stand, nicht der beim Absenden
            }
            if (g_callback) {
                const OverlayHandle handle = g_press.handle;  // Kopie: der Rückruf baut das Overlay neu auf
                const bool shift = g_press.shift;
                if (gesture == OverlayGesture::DragEnd || gesture == OverlayGesture::DragCancel) g_press.posted = false;
                g_callback(handle, gesture, gesture == OverlayGesture::Click ? 0 : pixels, shift);
            }
            return 0;
        }
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

HWND ensureWindow(HWND owner) {
    if (g_window && GetWindow(g_window, GW_OWNER) == owner) return g_window;
    if (g_window) DestroyWindow(g_window);
    static bool registered = false;
    const auto module = reinterpret_cast<HINSTANCE>(&__ImageBase);
    if (!registered) {
        WNDCLASSEXW wc{sizeof(wc)};
        wc.lpfnWndProc = windowProc;
        wc.hInstance = module;
        wc.lpszClassName = kClass;
        RegisterClassExW(&wc);
        registered = true;
    }
    g_window = CreateWindowExW(WS_EX_LAYERED | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW, kClass, L"", WS_POPUP, 0, 0, 1, 1, owner, nullptr, module,
                               nullptr);
    return g_window;
}

// Vorbelegtes (premultiplied) BGRA-Pixel
std::uint32_t pixel(COLORREF color, double alpha) {
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
void drawText(Canvas& canvas, HDC screen, HFONT font, const std::wstring& text, RECT rect, UINT format, COLORREF color, double alpha) {
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

}  // namespace

void showOverlay(HWND owner, const OverlayContent& content, OverlayCallback callback) {
    HWND hwnd = ensureWindow(owner);
    if (!hwnd) return;
    g_content = content;
    g_callback = callback;
    g_displaced = false;

    const int m = margin(owner), glow = scaled(owner, 9);
    const RECT& e = content.element;
    const int width = std::max<int>(e.right - e.left, 1) + 2 * m, height = std::max<int>(e.bottom - e.top, 1) + 2 * m;
    g_origin = {e.left - m, e.top - m};

    BITMAPINFO info{};
    info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    info.bmiHeader.biWidth = width;
    info.bmiHeader.biHeight = -height;
    info.bmiHeader.biPlanes = 1;
    info.bmiHeader.biBitCount = 32;
    info.bmiHeader.biCompression = BI_RGB;
    void* bits = nullptr;
    HDC screen = GetDC(nullptr);
    HDC dc = CreateCompatibleDC(screen);
    HBITMAP bitmap = CreateDIBSection(dc, &info, DIB_RGB_COLORS, &bits, nullptr, 0);
    if (!bitmap) {
        DeleteDC(dc);
        ReleaseDC(nullptr, screen);
        return;
    }
    HGDIOBJ previous = SelectObject(dc, bitmap);
    auto* px = static_cast<std::uint32_t*>(bits);
    std::fill(px, px + static_cast<size_t>(width) * height, 0u);
    Canvas canvas(px, width, height);

    // Leuchtrahmen: feine Kante am Element, weich nach außen auslaufend
    const int left = m, top = m, right = width - m, bottom = height - m;
    for (int y = top - glow; y < bottom + glow; ++y) {
        for (int x = left - glow; x < right + glow; ++x) {
            const int dx = x < left ? left - x : (x >= right ? x - right + 1 : 0);
            const int dy = y < top ? top - y : (y >= bottom ? y - bottom + 1 : 0);
            double alpha = 0;
            if (dx == 0 && dy == 0) {
                const int inside = std::min({x - left, right - 1 - x, y - top, bottom - 1 - y});
                if (inside < 1) alpha = 0.9;
                else if (inside < 4) alpha = 0.12 * (4 - inside) / 3.0;
            } else {
                const double d = std::sqrt(static_cast<double>(dx * dx + dy * dy));
                if (d <= glow) alpha = 0.45 * std::pow(1.0 - d / glow, 1.7);
            }
            canvas.blend(x, y, kBlue, alpha);
        }
    }

    // Winziges, unaufdringliches Etikett über dem Rahmen
    HFONT font = CreateFontW(-scaled(owner, 9), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                             ANTIALIASED_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    drawText(canvas, screen, font, L"rethink chart", {left, top - glow - scaled(owner, 13), width, top - glow}, DT_LEFT | DT_BOTTOM, kBlue, 0.85);

    // Griffe und Datenblatt-Symbol (Fensterkoordinaten)
    for (auto& h : content.handles) {
        RECT r = h.rect;
        OffsetRect(&r, -g_origin.x, -g_origin.y);
        switch (h.style) {
            case HandleStyle::Datasheet: {
                canvas.fill(r, kBlue, 1.0);
                // stilisiertes Tabellenraster
                const int w = r.right - r.left, hh = r.bottom - r.top;
                for (int i = 1; i <= 2; ++i) {
                    const int x = r.left + w * i / 3, y = r.top + hh * i / 3;
                    canvas.fill({x, r.top + 3, x + 1, r.bottom - 3}, RGB(255, 255, 255), 0.95);
                    canvas.fill({r.left + 3, y, r.right - 3, y + 1}, RGB(255, 255, 255), 0.95);
                }
                break;
            }
            case HandleStyle::Triangle:
            case HandleStyle::Arrow: canvas.triangle(r, h.pointing); break;
            case HandleStyle::Grip: canvas.grip(r); break;
            case HandleStyle::Hidden: canvas.fill(r, kBlue, 0.004); break;  // fast unsichtbar, aber für das Fenster vorhanden
        }
    }

    // Kurzinfo beim Ziehen: weißes Kästchen mit blauem Rand über dem Element
    if (!content.tooltip.empty()) {
        HFONT tipFont = CreateFontW(-scaled(owner, 11), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                                    CLIP_DEFAULT_PRECIS, ANTIALIASED_QUALITY, DEFAULT_PITCH, L"Segoe UI");
        HDC measure = CreateCompatibleDC(screen);
        HGDIOBJ old = SelectObject(measure, tipFont);
        SIZE extent{};
        GetTextExtentPoint32W(measure, content.tooltip.c_str(), static_cast<int>(content.tooltip.size()), &extent);
        SelectObject(measure, old);
        DeleteDC(measure);
        const int padX = scaled(owner, 6), padY = scaled(owner, 3);
        const int tw = extent.cx + 2 * padX, th = extent.cy + 2 * padY;
        int x = content.tooltipAt.x - g_origin.x - tw / 2, y = content.tooltipAt.y - g_origin.y - th;
        x = std::clamp(x, 0, std::max(0, width - tw));
        y = std::clamp(y, 0, std::max(0, height - th));
        canvas.fill({x, y, x + tw, y + th}, kBlue, 1.0);
        canvas.fill({x + 1, y + 1, x + tw - 1, y + th - 1}, RGB(255, 255, 255), 1.0);
        drawText(canvas, screen, tipFont, content.tooltip, {x + padX, y + padY, x + tw - padX, y + th - padY}, DT_LEFT | DT_TOP, RGB(32, 32, 32), 1.0);
        DeleteObject(tipFont);
    }
    DeleteObject(font);

    POINT position = g_origin, source{0, 0};
    SIZE size{width, height};
    BLENDFUNCTION blend{AC_SRC_OVER, 0, 255, AC_SRC_ALPHA};
    UpdateLayeredWindow(hwnd, screen, &position, &size, dc, &source, 0, &blend, ULW_ALPHA);
    ShowWindow(hwnd, SW_SHOWNOACTIVATE);

    SelectObject(dc, previous);
    DeleteObject(bitmap);
    DeleteDC(dc);
    ReleaseDC(nullptr, screen);
}

void hideOverlay() {
    if (g_window) {
        if (g_press.active) release(g_window, OverlayGesture::DragCancel);
        ShowWindow(g_window, SW_HIDE);
    }
    g_content = {};
}

bool overlayVisible() { return g_window && IsWindowVisible(g_window); }

HWND overlayWindow() { return g_window; }

void offsetOverlay(int dx, int dy) {
    if (!overlayVisible()) return;
    SetWindowPos(g_window, nullptr, g_origin.x + dx, g_origin.y + dy, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
    g_displaced = dx != 0 || dy != 0;
}

bool overlayDisplaced() { return g_displaced; }

RECT overlayRect() {
    RECT r{};
    if (g_window) GetWindowRect(g_window, &r);
    return r;
}

}  // namespace rethink::addin
