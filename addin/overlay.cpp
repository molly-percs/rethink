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
            const bool valueHandle = index >= 0 && g_content.handles[index].action != OverlayAction::Datasheet;
            SetCursor(LoadCursorW(nullptr, valueHandle ? IDC_SIZENS : IDC_HAND));
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
            // ▲▼: Klick oder Ziehen entscheidet sich erst an Bewegung bzw. Loslassen
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
            const int pixels = g_press.start.y - p.y;
            if (!g_press.dragging && std::abs(pixels) < GetSystemMetrics(SM_CYDRAG)) return 0;
            g_press.dragging = true;
            g_press.pixels = pixels;
            SetCursor(LoadCursorW(nullptr, IDC_SIZENS));
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
    // Dreieck mit Spitze nach oben (up) oder unten, weiße Kontur für Sichtbarkeit auf Segmenten
    void triangle(const RECT& r, bool up) {
        const double cx = (r.left + r.right) / 2.0, half = (r.right - r.left) / 2.0;
        for (int y = r.top; y < r.bottom; ++y) {
            const double t = (y + 0.5 - r.top) / (r.bottom - r.top);  // 0 oben … 1 unten
            const double w = half * (up ? t : 1 - t);
            for (int x = r.left; x < r.right; ++x) {
                const double d = std::fabs(x + 0.5 - cx);
                if (d <= w - 1.0) blend(x, y, kBlue, 1.0);
                else if (d <= w + 0.6) blend(x, y, RGB(255, 255, 255), 0.95);
            }
        }
    }

private:
    std::uint32_t* px_;
    int w_, h_;
};

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
    {
        // GDI-Text in eine eigene Maske zeichnen und als Alpha übernehmen
        const wchar_t* label = L"rethink chart";
        HDC maskDc = CreateCompatibleDC(screen);
        BITMAPINFO maskInfo = info;
        void* maskBits = nullptr;
        HBITMAP mask = CreateDIBSection(maskDc, &maskInfo, DIB_RGB_COLORS, &maskBits, nullptr, 0);
        if (mask) {
            HGDIOBJ oldMask = SelectObject(maskDc, mask);
            HGDIOBJ oldFont = SelectObject(maskDc, font);
            SetBkMode(maskDc, TRANSPARENT);
            SetTextColor(maskDc, RGB(255, 255, 255));
            RECT textRect{left, top - glow - scaled(owner, 13), width, top - glow};
            DrawTextW(maskDc, label, -1, &textRect, DT_LEFT | DT_BOTTOM | DT_SINGLELINE | DT_NOPREFIX);
            const auto* mp = static_cast<const std::uint32_t*>(maskBits);
            for (int i = 0; i < width * height; ++i) {
                const double coverage = (mp[i] & 0xFF) / 255.0;
                if (coverage > 0) canvas.blend(i % width, i / width, kBlue, 0.85 * coverage);
            }
            SelectObject(maskDc, oldFont);
            SelectObject(maskDc, oldMask);
            DeleteObject(mask);
        }
        DeleteDC(maskDc);
    }
    DeleteObject(font);

    // Griffe und Datenblatt-Symbol (Fensterkoordinaten)
    for (auto& h : content.handles) {
        RECT r = h.rect;
        OffsetRect(&r, -g_origin.x, -g_origin.y);
        if (h.action == OverlayAction::Datasheet) {
            canvas.fill(r, kBlue, 1.0);
            // stilisiertes Tabellenraster
            const int w = r.right - r.left, hh = r.bottom - r.top;
            for (int i = 1; i <= 2; ++i) {
                const int x = r.left + w * i / 3, y = r.top + hh * i / 3;
                canvas.fill({x, r.top + 3, x + 1, r.bottom - 3}, RGB(255, 255, 255), 0.95);
                canvas.fill({r.left + 3, y, r.right - 3, y + 1}, RGB(255, 255, 255), 0.95);
            }
        } else {
            canvas.triangle(r, h.action == OverlayAction::Up);
        }
    }

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
