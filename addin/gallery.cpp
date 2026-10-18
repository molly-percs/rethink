#include "gallery.h"

#include <windowsx.h>

#include <algorithm>
#include <array>
#include <optional>
#include <vector>

extern "C" IMAGE_DOS_HEADER __ImageBase;

namespace rethink::addin {
namespace {

using engine::Orientation;

constexpr wchar_t kClass[] = L"rethinkGallery";
constexpr COLORREF kBlue = RGB(37, 120, 235), kHover = RGB(232, 240, 253), kText = RGB(40, 40, 40), kGrey = RGB(160, 160, 160);
constexpr COLORREF kC1 = RGB(31, 78, 121), kC2 = RGB(46, 117, 182), kC3 = RGB(157, 195, 230);
constexpr COLORREF kD1 = RGB(190, 190, 190), kD2 = RGB(210, 210, 210), kD3 = RGB(230, 230, 230);

enum class Icon { Stacked, Clustered, Percent, Line, Area, MekkoPercent, MekkoUnits, Pie, Scatter, Bubble, Waterfall, WaterfallDown, Gantt };

struct Tile {
    const wchar_t* label;
    Icon icon;
    std::optional<ElementKind> kind;  // leer = noch nicht umgesetzt (ausgegraut)
    bool orientable = false;
};

const std::array<Tile, 13> kTiles = {{
    {L"Gestapelt", Icon::Stacked, std::nullopt, true},
    {L"Gruppiert", Icon::Clustered, std::nullopt, true},
    {L"100 %", Icon::Percent, std::nullopt, true},
    {L"Linie", Icon::Line, std::nullopt},
    {L"Fläche", Icon::Area, std::nullopt},
    {L"Mekko %", Icon::MekkoPercent, std::nullopt},
    {L"Mekko Einheiten", Icon::MekkoUnits, std::nullopt},
    {L"Kreis", Icon::Pie, std::nullopt},
    {L"Punkte", Icon::Scatter, std::nullopt},
    {L"Blasen", Icon::Bubble, std::nullopt},
    {L"Wasserfall", Icon::Waterfall, ElementKind::Waterfall, true},
    {L"Wasserfall abbauend", Icon::WaterfallDown, ElementKind::WaterfallDown, true},
    {L"Gantt / Zeitplan", Icon::Gantt, ElementKind::Gantt},
}};
constexpr int kColumns = 5;

struct Hit {
    int tile = -1;
    std::optional<Orientation> arrow;  // Pfeil um das Symbol
};

struct State {
    HWND hwnd = nullptr;
    HFONT font = nullptr;
    int dpi = 96;
    Hit hover;
    std::optional<GalleryChoice> choice;
    std::function<void(GalleryChoice)> chosen;
};
State g;

int s(int px) { return MulDiv(px, g.dpi, 96); }
int tileW() { return s(108); }
int tileH() { return s(96); }
int statusH() { return s(24); }

RECT tileRect(int i) {
    const int x = s(8) + (i % kColumns) * tileW(), y = s(8) + (i / kColumns) * tileH();
    return {x, y, x + tileW(), y + tileH()};
}

RECT iconRect(int i) {
    const RECT t = tileRect(i);
    const int w = s(56), h = s(40), x = (t.left + t.right - w) / 2, y = t.top + s(16);
    return {x, y, x + w, y + h};
}

// Pfeile um das Symbol: oben, unten, links, rechts
RECT arrowRect(int i, Orientation o) {
    const RECT r = iconRect(i);
    const int a = s(12), b = s(9), cx = (r.left + r.right) / 2, cy = (r.top + r.bottom) / 2;
    switch (o) {
        case Orientation::Up: return {cx - a / 2, r.top - b - s(3), cx + a / 2, r.top - s(3)};
        case Orientation::Down: return {cx - a / 2, r.bottom + s(3), cx + a / 2, r.bottom + b + s(3)};
        case Orientation::Left: return {r.left - b - s(4), cy - a / 2, r.left - s(4), cy + a / 2};
        case Orientation::Right: return {r.right + s(4), cy - a / 2, r.right + b + s(4), cy + a / 2};
    }
    return r;
}

Hit hitTest(int x, int y) {
    const POINT p{x, y};
    for (int i = 0; i < static_cast<int>(kTiles.size()); ++i) {
        const RECT t = tileRect(i);
        if (!PtInRect(&t, p)) continue;
        Hit h{i, std::nullopt};
        if (kTiles[i].orientable)
            for (Orientation o : {Orientation::Up, Orientation::Down, Orientation::Left, Orientation::Right}) {
                RECT a = arrowRect(i, o);
                InflateRect(&a, s(3), s(3));
                if (PtInRect(&a, p)) h.arrow = o;
            }
        return h;
    }
    return {};
}

void fillRect(HDC dc, int l, int t, int r, int b, COLORREF c) {
    HBRUSH brush = CreateSolidBrush(c);
    RECT rect{l, t, r, b};
    FillRect(dc, &rect, brush);
    DeleteObject(brush);
}

void triangle(HDC dc, const RECT& r, Orientation o, COLORREF c) {
    POINT p[3];
    const int cx = (r.left + r.right) / 2, cy = (r.top + r.bottom) / 2;
    switch (o) {
        case Orientation::Up: p[0] = {cx, r.top}, p[1] = {r.right, r.bottom}, p[2] = {r.left, r.bottom}; break;
        case Orientation::Down: p[0] = {r.left, r.top}, p[1] = {r.right, r.top}, p[2] = {cx, r.bottom}; break;
        case Orientation::Left: p[0] = {r.left, cy}, p[1] = {r.right, r.top}, p[2] = {r.right, r.bottom}; break;
        case Orientation::Right: p[0] = {r.left, r.top}, p[1] = {r.right, cy}, p[2] = {r.left, r.bottom}; break;
    }
    HBRUSH brush = CreateSolidBrush(c);
    HPEN pen = CreatePen(PS_SOLID, 1, c);
    HGDIOBJ ob = SelectObject(dc, brush), op = SelectObject(dc, pen);
    Polygon(dc, p, 3);
    SelectObject(dc, ob), SelectObject(dc, op);
    DeleteObject(brush), DeleteObject(pen);
}

// Schematisches Symbol eines Diagrammtyps (eigene Zeichnung, keine fremden Symbole)
void drawIcon(HDC dc, Icon icon, const RECT& r, bool enabled) {
    const COLORREF c1 = enabled ? kC1 : kD1, c2 = enabled ? kC2 : kD2, c3 = enabled ? kC3 : kD3;
    const int x = r.left, y = r.top, w = r.right - r.left, h = r.bottom - r.top;
    auto bar = [&](double fx, double fy, double fw, double fh, COLORREF c) {
        fillRect(dc, x + static_cast<int>(fx * w), y + static_cast<int>(fy * h), x + static_cast<int>((fx + fw) * w), y + static_cast<int>((fy + fh) * h), c);
    };
    HPEN axis = CreatePen(PS_SOLID, 1, enabled ? RGB(90, 90, 90) : kD1);
    HGDIOBJ oldPen = SelectObject(dc, axis);
    switch (icon) {
        case Icon::Stacked:
            for (int i = 0; i < 3; ++i) {
                const double fx = 0.08 + i * 0.32, top = 0.45 - i * 0.15;
                bar(fx, top, 0.22, 0.35 - top * 0.3, c3), bar(fx, top + 0.35 - top * 0.3, 0.22, 0.3, c2), bar(fx, 0.65 + 0.05, 0.22, 0.3, c1);
            }
            break;
        case Icon::Clustered:
            for (int i = 0; i < 3; ++i) bar(0.06 + i * 0.32, 0.5 - i * 0.12, 0.12, 0.5 + i * 0.12, c1), bar(0.19 + i * 0.32, 0.3 + i * 0.05, 0.12, 0.7 - i * 0.05, c3);
            break;
        case Icon::Percent:
            for (int i = 0; i < 3; ++i) {
                const double fx = 0.08 + i * 0.32, a = 0.3 + i * 0.1, b = 0.3 - i * 0.05;
                bar(fx, 0, 0.22, a, c3), bar(fx, a, 0.22, b, c2), bar(fx, a + b, 0.22, 1 - a - b, c1);
            }
            break;
        case Icon::Line: {
            HPEN pen = CreatePen(PS_SOLID, s(2), c2);
            SelectObject(dc, pen);
            const POINT pts[] = {{x, y + h * 8 / 10}, {x + w / 4, y + h / 2}, {x + w / 2, y + h * 6 / 10}, {x + w * 3 / 4, y + h / 4}, {x + w, y + h / 10}};
            Polyline(dc, pts, 5);
            SelectObject(dc, axis);
            DeleteObject(pen);
            for (auto& p : pts) fillRect(dc, p.x - s(2), p.y - s(2), p.x + s(2), p.y + s(2), c1);
            break;
        }
        case Icon::Area: {
            HBRUSH brush = CreateSolidBrush(c3);
            HPEN pen = CreatePen(PS_SOLID, 1, c2);
            HGDIOBJ ob = SelectObject(dc, brush);
            SelectObject(dc, pen);
            const POINT pts[] = {{x, y + h}, {x, y + h * 6 / 10}, {x + w / 3, y + h * 3 / 10}, {x + w * 2 / 3, y + h / 2}, {x + w, y + h / 5}, {x + w, y + h}};
            Polygon(dc, pts, 6);
            SelectObject(dc, ob), SelectObject(dc, axis);
            DeleteObject(brush), DeleteObject(pen);
            break;
        }
        case Icon::MekkoPercent:
            bar(0, 0, 0.3, 0.4, c3), bar(0, 0.4, 0.3, 0.6, c1);
            bar(0.31, 0, 0.45, 0.6, c3), bar(0.31, 0.6, 0.45, 0.4, c1);
            bar(0.77, 0, 0.23, 0.3, c3), bar(0.77, 0.3, 0.23, 0.7, c1);
            break;
        case Icon::MekkoUnits:
            bar(0, 0.3, 0.3, 0.3, c3), bar(0, 0.6, 0.3, 0.4, c1);
            bar(0.31, 0, 0.45, 0.5, c3), bar(0.31, 0.5, 0.45, 0.5, c1);
            bar(0.77, 0.55, 0.23, 0.2, c3), bar(0.77, 0.75, 0.23, 0.25, c1);
            break;
        case Icon::Pie: {
            const int d = std::min(w, h), cx = x + w / 2, cy = y + h / 2;
            HBRUSH b1 = CreateSolidBrush(c1), b3 = CreateSolidBrush(c3);
            HGDIOBJ ob = SelectObject(dc, b3);
            Ellipse(dc, cx - d / 2, cy - d / 2, cx + d / 2, cy + d / 2);
            SelectObject(dc, b1);
            Pie(dc, cx - d / 2, cy - d / 2, cx + d / 2, cy + d / 2, cx + d / 2, cy, cx, cy - d / 2);
            SelectObject(dc, ob);
            DeleteObject(b1), DeleteObject(b3);
            break;
        }
        case Icon::Scatter:
            for (auto [fx, fy] : {std::pair{0.1, 0.8}, {0.3, 0.55}, {0.45, 0.65}, {0.6, 0.35}, {0.8, 0.25}, {0.9, 0.1}})
                fillRect(dc, x + static_cast<int>(fx * w) - s(2), y + static_cast<int>(fy * h) - s(2), x + static_cast<int>(fx * w) + s(3),
                         y + static_cast<int>(fy * h) + s(3), c1);
            break;
        case Icon::Bubble: {
            HBRUSH b2 = CreateSolidBrush(c3);
            HGDIOBJ ob = SelectObject(dc, b2);
            HPEN pen = CreatePen(PS_SOLID, 1, c2);
            SelectObject(dc, pen);
            Ellipse(dc, x + w / 10, y + h / 2, x + w / 10 + h / 3, y + h / 2 + h / 3);
            Ellipse(dc, x + w * 4 / 10, y + h / 6, x + w * 4 / 10 + h / 2, y + h / 6 + h / 2);
            Ellipse(dc, x + w * 3 / 4, y + h / 2, x + w * 3 / 4 + h / 4, y + h / 2 + h / 4);
            SelectObject(dc, ob), SelectObject(dc, axis);
            DeleteObject(b2), DeleteObject(pen);
            break;
        }
        case Icon::Waterfall:
            bar(0.0, 0.5, 0.15, 0.5, c1), bar(0.21, 0.25, 0.15, 0.25, c2), bar(0.42, 0.25, 0.15, 0.15, c3);
            bar(0.63, 0.4, 0.15, 0.6, c1), bar(0.84, 0.15, 0.16, 0.85, c1);
            break;
        case Icon::WaterfallDown:
            bar(0.0, 0.0, 0.15, 1.0, c1), bar(0.21, 0.0, 0.15, 0.3, c2), bar(0.42, 0.3, 0.15, 0.25, c2), bar(0.63, 0.55, 0.15, 0.2, c3),
                bar(0.84, 0.75, 0.16, 0.25, c1);
            break;
        case Icon::Gantt: {
            bar(0.0, 0.05, 1.0, 0.12, c3);
            bar(0.05, 0.28, 0.4, 0.14, c1), bar(0.35, 0.5, 0.35, 0.14, c1), bar(0.6, 0.72, 0.3, 0.14, c2);
            HBRUSH b = CreateSolidBrush(enabled ? RGB(237, 125, 49) : kD1);
            HGDIOBJ ob = SelectObject(dc, b);
            const int cx = x + w * 95 / 100, cy = y + h * 79 / 100, d = s(5);
            const POINT dia[] = {{cx, cy - d}, {cx + d, cy}, {cx, cy + d}, {cx - d, cy}};
            Polygon(dc, dia, 4);
            SelectObject(dc, ob);
            DeleteObject(b);
            break;
        }
    }
    if (icon != Icon::Pie && icon != Icon::Scatter && icon != Icon::Bubble && icon != Icon::Gantt) {
        MoveToEx(dc, x - s(2), y + h, nullptr);
        LineTo(dc, x + w + s(2), y + h);
    }
    SelectObject(dc, oldPen);
    DeleteObject(axis);
}

void paint(HWND hwnd, HDC target) {
    RECT client;
    GetClientRect(hwnd, &client);
    HDC dc = CreateCompatibleDC(target);
    HBITMAP bmp = CreateCompatibleBitmap(target, client.right, client.bottom);
    HGDIOBJ oldBmp = SelectObject(dc, bmp), oldFont = SelectObject(dc, g.font);
    SetBkMode(dc, TRANSPARENT);
    fillRect(dc, 0, 0, client.right, client.bottom, RGB(255, 255, 255));
    for (int i = 0; i < static_cast<int>(kTiles.size()); ++i) {
        const Tile& t = kTiles[i];
        const bool enabled = t.kind.has_value(), hovered = g.hover.tile == i;
        const RECT r = tileRect(i);
        if (hovered && enabled) fillRect(dc, r.left + 1, r.top + 1, r.right - 1, r.bottom - 1, kHover);
        drawIcon(dc, t.icon, iconRect(i), enabled);
        if (hovered && enabled && t.orientable)
            for (Orientation o : {Orientation::Up, Orientation::Down, Orientation::Left, Orientation::Right})
                triangle(dc, arrowRect(i, o), o, g.hover.arrow == o ? kBlue : RGB(120, 160, 220));
        RECT label{r.left + s(4), iconRect(i).bottom + s(12), r.right - s(4), r.bottom - s(2)};
        SetTextColor(dc, enabled ? kText : kGrey);
        DrawTextW(dc, t.label, -1, &label, DT_CENTER | DT_TOP | DT_WORDBREAK | DT_NOPREFIX);
    }
    RECT status{s(10), client.bottom - statusH(), client.right - s(10), client.bottom};
    fillRect(dc, 0, client.bottom - statusH(), client.right, client.bottom, RGB(243, 243, 243));
    SetTextColor(dc, RGB(100, 100, 100));
    DrawTextW(dc, L"Pfeile um das Symbol wählen die Ausrichtung · grau: noch nicht in rethink umgesetzt", -1, &status,
              DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
    BitBlt(target, 0, 0, client.right, client.bottom, dc, 0, 0, SRCCOPY);
    SelectObject(dc, oldFont), SelectObject(dc, oldBmp);
    DeleteObject(bmp);
    DeleteDC(dc);
}

LRESULT CALLBACK windowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_PAINT: {
            PAINTSTRUCT ps;
            HDC dc = BeginPaint(hwnd, &ps);
            paint(hwnd, dc);
            EndPaint(hwnd, &ps);
            return 0;
        }
        case WM_ERASEBKGND: return 1;
        case WM_MOUSEMOVE: {
            const Hit h = hitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
            if (h.tile != g.hover.tile || h.arrow != g.hover.arrow) {
                g.hover = h;
                InvalidateRect(hwnd, nullptr, FALSE);
            }
            TRACKMOUSEEVENT tme{sizeof(tme), TME_LEAVE, hwnd, 0};
            TrackMouseEvent(&tme);
            return 0;
        }
        case WM_MOUSELEAVE:
            g.hover = {};
            InvalidateRect(hwnd, nullptr, FALSE);
            return 0;
        case WM_SETCURSOR: {
            POINT p;
            GetCursorPos(&p);
            ScreenToClient(hwnd, &p);
            const Hit h = hitTest(p.x, p.y);
            SetCursor(LoadCursorW(nullptr, h.tile >= 0 && kTiles[h.tile].kind ? IDC_HAND : IDC_ARROW));
            return TRUE;
        }
        case WM_LBUTTONUP: {
            const Hit h = hitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
            if (h.tile < 0 || !kTiles[h.tile].kind) return 0;  // ausgegraut: nichts
            g.choice = GalleryChoice{*kTiles[h.tile].kind, h.arrow.value_or(Orientation::Up)};
            DestroyWindow(hwnd);
            return 0;
        }
        case WM_KEYDOWN:
            if (wParam == VK_ESCAPE) DestroyWindow(hwnd);
            return 0;
        case WM_ACTIVATE:
            if (LOWORD(wParam) == WA_INACTIVE) PostMessageW(hwnd, WM_CLOSE, 0, 0);  // Klick daneben
            return 0;
        case WM_CLOSE: DestroyWindow(hwnd); return 0;
        case WM_NCDESTROY: {
            auto chosen = std::move(g.chosen);
            const auto choice = g.choice;
            if (g.font) DeleteObject(g.font);
            g = {};
            if (choice && chosen) chosen(*choice);  // erst nach dem Schließen einfügen
            return 0;
        }
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

}  // namespace

void openGallery(HWND owner, POINT at, std::function<void(GalleryChoice)> chosen) {
    if (g.hwnd) DestroyWindow(g.hwnd);
    const auto module = reinterpret_cast<HINSTANCE>(&__ImageBase);
    static bool registered = false;
    if (!registered) {
        WNDCLASSEXW wc{sizeof(wc)};
        wc.style = CS_DROPSHADOW;
        wc.lpfnWndProc = windowProc;
        wc.hInstance = module;
        wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kClass;
        RegisterClassExW(&wc);
        registered = true;
    }
    g.chosen = std::move(chosen);
    const UINT dpi = GetDpiForWindow(owner);
    g.dpi = dpi ? static_cast<int>(dpi) : 96;
    g.font = CreateFontW(-MulDiv(9, g.dpi, 72), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                         CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    const int rows = (static_cast<int>(kTiles.size()) + kColumns - 1) / kColumns;
    const int w = s(16) + kColumns * tileW() + 2, h = s(16) + rows * tileH() + statusH() + 2;
    MONITORINFO mi{sizeof(mi)};
    GetMonitorInfoW(MonitorFromPoint(at, MONITOR_DEFAULTTONEAREST), &mi);
    const int x = std::clamp<int>(at.x - s(20), mi.rcWork.left, mi.rcWork.right - w), y = std::clamp<int>(at.y + s(12), mi.rcWork.top, mi.rcWork.bottom - h);
    g.hwnd = CreateWindowExW(WS_EX_TOOLWINDOW, kClass, L"rethink – Diagramm einfügen", WS_POPUP | WS_BORDER, x, y, w, h, owner, nullptr, module, nullptr);
    if (!g.hwnd) return;
    ShowWindow(g.hwnd, SW_SHOW);
    SetForegroundWindow(g.hwnd);
    SetFocus(g.hwnd);
}

bool galleryOpen() { return g.hwnd != nullptr; }

}  // namespace rethink::addin
