#include "radial_menu.h"

#include "canvas.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <numbers>
#include <vector>

namespace rethink::addin {
namespace {

constexpr wchar_t kClass[] = L"rethinkGanttRadialMenu";
constexpr UINT_PTR kTimer = 1;
constexpr COLORREF kRed = RGB(220, 30, 30), kInk = RGB(40, 40, 40), kGrey = RGB(170, 170, 170);

struct State {
    HWND hwnd = nullptr;
    POINT center{};  // Bildschirm
    RECT plot{};
    POINT origin{};  // Bildschirmposition der linken oberen Fensterecke
    SIZE size{};
    std::vector<RadialButton> buttons;
    RadialCallback callback = nullptr;
    int hot = -1;  // -1 = keine, buttons.size() = Mitte (schließen)
};
State g;

int px(int v) { return MulDiv(v, static_cast<int>(GetDpiForSystem()), 96); }
int ringRadius() { return px(46); }
int buttonRadius() { return px(16); }

// Mittelpunkt der Schaltfläche i (Bildschirm), im Uhrzeigersinn ab oben
POINT buttonCenter(size_t i) {
    const double angle = -std::numbers::pi / 2 + 2 * std::numbers::pi * static_cast<double>(i) / static_cast<double>(g.buttons.size());
    return {g.center.x + static_cast<LONG>(std::lround(ringRadius() * std::cos(angle))),
            g.center.y + static_cast<LONG>(std::lround(ringRadius() * std::sin(angle)))};
}

int buttonAt(POINT screen) {
    auto inside = [&](POINT c, int r) { return std::hypot(screen.x - c.x, screen.y - c.y) <= r; };
    for (size_t i = 0; i < g.buttons.size(); ++i)
        if (inside(buttonCenter(i), buttonRadius())) return static_cast<int>(i);
    if (inside(g.center, px(11))) return static_cast<int>(g.buttons.size());
    return -1;
}

// Kreisscheibe mit weicher Kante
void disc(Canvas& c, double cx, double cy, double r, COLORREF color, double alpha) {
    for (int y = static_cast<int>(cy - r - 1); y <= static_cast<int>(cy + r + 1); ++y)
        for (int x = static_cast<int>(cx - r - 1); x <= static_cast<int>(cx + r + 1); ++x) {
            const double d = std::hypot(x + 0.5 - cx, y + 0.5 - cy);
            c.blend(x, y, color, alpha * std::clamp(r - d + 0.5, 0.0, 1.0));
        }
}

void ring(Canvas& c, double cx, double cy, double r, double width, COLORREF color, double alpha) {
    for (int y = static_cast<int>(cy - r - 2); y <= static_cast<int>(cy + r + 2); ++y)
        for (int x = static_cast<int>(cx - r - 2); x <= static_cast<int>(cx + r + 2); ++x) {
            const double d = std::fabs(std::hypot(x + 0.5 - cx, y + 0.5 - cy) - r);
            c.blend(x, y, color, alpha * std::clamp(width / 2 - d + 0.5, 0.0, 1.0));
        }
}

// Linie mit weicher Kante (Abstand zum Segment)
void line(Canvas& c, double x0, double y0, double x1, double y1, double width, COLORREF color, double alpha) {
    const double dx = x1 - x0, dy = y1 - y0, len2 = std::max(dx * dx + dy * dy, 1e-9);
    for (int y = static_cast<int>(std::min(y0, y1) - width - 1); y <= static_cast<int>(std::max(y0, y1) + width + 1); ++y)
        for (int x = static_cast<int>(std::min(x0, x1) - width - 1); x <= static_cast<int>(std::max(x0, x1) + width + 1); ++x) {
            const double t = std::clamp(((x + 0.5 - x0) * dx + (y + 0.5 - y0) * dy) / len2, 0.0, 1.0);
            const double d = std::hypot(x + 0.5 - (x0 + t * dx), y + 0.5 - (y0 + t * dy));
            c.blend(x, y, color, alpha * std::clamp(width / 2 - d + 0.5, 0.0, 1.0));
        }
}

void polygon(Canvas& c, const std::vector<POINT>& points, COLORREF color, double alpha) {
    LONG top = points[0].y, bottom = points[0].y, left = points[0].x, right = points[0].x;
    for (auto& p : points) top = std::min(top, p.y), bottom = std::max(bottom, p.y), left = std::min(left, p.x), right = std::max(right, p.x);
    for (LONG y = top; y < bottom; ++y)
        for (LONG x = left; x < right; ++x) {
            bool in = false;  // Strahltest
            for (size_t i = 0, j = points.size() - 1; i < points.size(); j = i++) {
                const POINT &a = points[i], &b = points[j];
                if ((a.y > y) != (b.y > y) && x + 0.5 < (b.x - a.x) * (y + 0.5 - a.y) / static_cast<double>(b.y - a.y) + a.x) in = !in;
            }
            if (in) c.blend(x, y, color, alpha);
        }
}

// Symbol des Befehls in der Schaltfläche (Fensterkoordinaten)
void icon(Canvas& c, const std::wstring& command, int cx, int cy, COLORREF ink) {
    const int u = px(5);
    if (command == L"new=bar") {
        c.fill({cx - 2 * u, cy - u, cx + 2 * u, cy + u}, ink, 1.0);
    } else if (command == L"new=pentagon") {
        polygon(c, {{cx - 2 * u, cy - u}, {cx + u, cy - u}, {cx + 2 * u, cy}, {cx + u, cy + u}, {cx - 2 * u, cy + u}}, ink, 1.0);
    } else if (command == L"new=milestone") {
        polygon(c, {{cx, cy - u - u / 2}, {cx + u + u / 2, cy + u}, {cx - u - u / 2, cy + u}}, ink, 1.0);
    } else if (command == L"new=bracket") {
        line(c, cx - 2 * u, cy, cx + 2 * u, cy, px(2), ink, 1.0);
        line(c, cx - 2 * u, cy - u, cx - 2 * u, cy + px(1), px(2), ink, 1.0);
        line(c, cx + 2 * u, cy - u, cx + 2 * u, cy + px(1), px(2), ink, 1.0);
    } else if (command == L"new=shade") {
        c.fill({cx - 2 * u, cy - 2 * u + px(2), cx + 2 * u, cy + 2 * u - px(2)}, ink, 0.3);
    } else {  // weitere Optionen: drei Punkte
        for (int k = -1; k <= 1; ++k) disc(c, cx + k * u, cy, px(2), ink, 1.0);
    }
}

void render() {
    const int w = g.size.cx, h = g.size.cy;
    BITMAPINFO info{};
    info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    info.bmiHeader.biWidth = w;
    info.bmiHeader.biHeight = -h;
    info.bmiHeader.biPlanes = 1;
    info.bmiHeader.biBitCount = 32;
    info.bmiHeader.biCompression = BI_RGB;
    void* bits = nullptr;
    HDC screen = GetDC(nullptr), dc = CreateCompatibleDC(screen);
    HBITMAP bitmap = CreateDIBSection(dc, &info, DIB_RGB_COLORS, &bits, nullptr, 0);
    if (!bitmap) {
        DeleteDC(dc);
        ReleaseDC(nullptr, screen);
        return;
    }
    HGDIOBJ previous = SelectObject(dc, bitmap);
    auto* pixels = static_cast<std::uint32_t*>(bits);
    std::fill(pixels, pixels + static_cast<size_t>(w) * h, 0u);
    Canvas c(pixels, w, h);
    const int cx = g.center.x - g.origin.x, cy = g.center.y - g.origin.y, R = ringRadius(), r = buttonRadius();

    // Fadenkreuz: wo das neue Element entsteht (außerhalb des Rings)
    const int gap = R + r + px(4);
    const int left = g.plot.left - g.origin.x, right = g.plot.right - g.origin.x, top = g.plot.top - g.origin.y, bottom = g.plot.bottom - g.origin.y;
    if (cx - gap > left) c.fill({left, cy, cx - gap, cy + 1}, kRed, 0.9);
    if (right > cx + gap) c.fill({cx + gap, cy, right, cy + 1}, kRed, 0.9);
    if (cy - gap > top) c.fill({cx, top, cx + 1, cy - gap}, kRed, 0.9);
    if (bottom > cy + gap) c.fill({cx, cy + gap, cx + 1, bottom}, kRed, 0.9);

    // Ring als Hintergrund (fängt die Maus zwischen den Schaltflächen ab)
    disc(c, cx, cy, R + r + px(3), RGB(255, 255, 255), 0.6);
    ring(c, cx, cy, R + r + px(3), 1.0, kGrey, 0.8);
    for (size_t i = 0; i < g.buttons.size(); ++i) {
        const POINT b = buttonCenter(i);
        const int bx = b.x - g.origin.x, by = b.y - g.origin.y;
        const bool hot = static_cast<int>(i) == g.hot && g.buttons[i].enabled;
        disc(c, bx, by, r, hot ? kBlue : RGB(255, 255, 255), 1.0);
        ring(c, bx, by, r, 1.2, hot ? kBlue : kGrey, 1.0);
        icon(c, g.buttons[i].command, bx, by, !g.buttons[i].enabled ? kGrey : hot ? RGB(255, 255, 255) : kInk);
    }
    // Mitte: schließen
    const bool hotClose = g.hot == static_cast<int>(g.buttons.size());
    disc(c, cx, cy, px(11), hotClose ? kRed : RGB(255, 255, 255), 1.0);
    ring(c, cx, cy, px(11), 1.2, hotClose ? kRed : kGrey, 1.0);
    const int k = px(4);
    line(c, cx - k, cy - k, cx + k, cy + k, px(2), hotClose ? RGB(255, 255, 255) : kInk, 1.0);
    line(c, cx - k, cy + k, cx + k, cy - k, px(2), hotClose ? RGB(255, 255, 255) : kInk, 1.0);

    // Kurzinfo der Schaltfläche unter dem Ring
    if (g.hot >= 0) {
        const std::wstring text = g.hot < static_cast<int>(g.buttons.size()) ? g.buttons[g.hot].label : L"Schließen";
        HFONT font = CreateFontW(-px(12), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                                 ANTIALIASED_QUALITY, DEFAULT_PITCH, L"Segoe UI");
        HGDIOBJ old = SelectObject(dc, font);
        SIZE extent{};
        GetTextExtentPoint32W(dc, text.c_str(), static_cast<int>(text.size()), &extent);
        SelectObject(dc, old);
        const int tw = extent.cx + px(12), th = extent.cy + px(6), tx = std::clamp(cx - tw / 2, 0, std::max(0, w - tw)), ty = cy + R + r + px(8);
        c.fill({tx, ty, tx + tw, ty + th}, kGrey, 1.0);
        c.fill({tx + 1, ty + 1, tx + tw - 1, ty + th - 1}, RGB(255, 255, 255), 1.0);
        drawText(c, screen, font, text, {tx + px(6), ty + px(3), tx + tw - px(6), ty + th - px(3)}, DT_LEFT | DT_TOP, kInk, 1.0);
        DeleteObject(font);
    }

    POINT position = g.origin, source{0, 0};
    BLENDFUNCTION blend{AC_SRC_OVER, 0, 255, AC_SRC_ALPHA};
    UpdateLayeredWindow(g.hwnd, screen, &position, &g.size, dc, &source, 0, &blend, ULW_ALPHA);
    SelectObject(dc, previous);
    DeleteObject(bitmap);
    DeleteDC(dc);
    ReleaseDC(nullptr, screen);
}

void choose(int index) {
    if (index < 0) return;
    if (index >= static_cast<int>(g.buttons.size())) return hideRadialMenu();
    if (!g.buttons[index].enabled) return;
    const std::wstring command = g.buttons[index].command;
    const RadialCallback callback = g.callback;
    hideRadialMenu();
    if (callback) callback(command);
}

LRESULT CALLBACK proc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
        case WM_MOUSEACTIVATE: return MA_NOACTIVATE;
        case WM_MOUSEMOVE: {
            POINT p{static_cast<short>(LOWORD(lParam)), static_cast<short>(HIWORD(lParam))};
            ClientToScreen(hwnd, &p);
            if (const int hot = buttonAt(p); hot != g.hot) {
                g.hot = hot;
                render();
            }
            return 0;
        }
        case WM_LBUTTONUP: {
            POINT p{static_cast<short>(LOWORD(lParam)), static_cast<short>(HIWORD(lParam))};
            ClientToScreen(hwnd, &p);
            choose(buttonAt(p));
            return 0;
        }
        case WM_TIMER: {
            // schließen bei Esc, einem Klick außerhalb des Rings oder wenn PowerPoint nicht mehr vorn ist
            POINT cursor{};
            GetCursorPos(&cursor);
            const bool outside = std::hypot(cursor.x - g.center.x, cursor.y - g.center.y) > ringRadius() + buttonRadius() + px(4);
            const bool pressed = (GetAsyncKeyState(VK_LBUTTON) | GetAsyncKeyState(VK_RBUTTON) | GetAsyncKeyState(VK_MBUTTON)) & 0x8000;
            DWORD process = 0;
            GetWindowThreadProcessId(GetForegroundWindow(), &process);
            if ((GetAsyncKeyState(VK_ESCAPE) & 0x8001) || (pressed && outside) || process != GetCurrentProcessId()) hideRadialMenu();
            return 0;
        }
        case WM_DESTROY: KillTimer(hwnd, kTimer); return 0;
    }
    return DefWindowProcW(hwnd, message, wParam, lParam);
}

}  // namespace

void showRadialMenu(HWND owner, POINT center, const RECT& plot, std::vector<RadialButton> buttons, RadialCallback callback) {
    hideRadialMenu();
    static const bool registered = [] {
        WNDCLASSW wc{};
        wc.lpfnWndProc = proc;
        wc.hInstance = GetModuleHandleW(nullptr);
        wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kClass;
        return RegisterClassW(&wc) != 0;
    }();
    if (!registered || buttons.empty()) return;
    g.center = center, g.plot = plot, g.buttons = std::move(buttons), g.callback = callback, g.hot = -1;
    // Fenster: Zeitbereich und Ring samt Kurzinfo darunter
    const int reach = ringRadius() + buttonRadius() + px(4);
    RECT area{center.x - std::max(reach, px(90)), center.y - reach, center.x + std::max(reach, px(90)), center.y + reach + px(36)};
    UnionRect(&area, &area, &plot);
    g.origin = {area.left, area.top};
    g.size = {area.right - area.left, area.bottom - area.top};
    g.hwnd = CreateWindowExW(WS_EX_LAYERED | WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_NOACTIVATE, kClass, L"", WS_POPUP, area.left, area.top, g.size.cx,
                             g.size.cy, owner, nullptr, GetModuleHandleW(nullptr), nullptr);
    if (!g.hwnd) return;
    render();
    ShowWindow(g.hwnd, SW_SHOWNOACTIVATE);
    SetTimer(g.hwnd, kTimer, 40, nullptr);
}

void hideRadialMenu() {
    if (!g.hwnd) return;
    HWND hwnd = g.hwnd;
    g.hwnd = nullptr;
    DestroyWindow(hwnd);
}

HWND radialMenuWindow() { return g.hwnd; }

std::wstring radialMenuText() {
    std::wstring text;
    if (!g.hwnd) return text;
    for (auto& b : g.buttons) text += b.label + L"=" + b.command + (b.enabled ? L"" : L" (aus)") + L"\n";
    return text;
}

bool clickRadialButton(const std::wstring& command) {
    for (size_t i = 0; g.hwnd && i < g.buttons.size(); ++i)
        if (g.buttons[i].command == command && g.buttons[i].enabled) {
            choose(static_cast<int>(i));
            return true;
        }
    return false;
}

}  // namespace rethink::addin
