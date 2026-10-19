#include "overlay.h"

#include "canvas.h"

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

HWND g_window = nullptr;
OverlayContent g_content;
POINT g_origin{};  // Bildschirmposition der linken oberen Fensterecke
bool g_displaced = false;
OverlayCallback g_callback = nullptr;
int g_hover = -1;  // Griff unter der Maus (Index in g_content.handles), -1 = keiner

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

bool sameHandle(const OverlayHandle& a, const OverlayHandle& b) {
    return a.action == b.action && a.row == b.row && a.column == b.column && a.role == b.role;
}

// Griff unter der Maus: beim Drücken der gedrückte, sonst der unter dem Zeiger (nur wenn das Overlay dort oben liegt)
int hoverTarget() {
    if (g_press.active) {
        for (size_t i = 0; i < g_content.handles.size(); ++i)
            if (sameHandle(g_content.handles[i], g_press.handle)) return static_cast<int>(i);
        return -1;
    }
    POINT p{};
    GetCursorPos(&p);
    return g_window && WindowFromPoint(p) == g_window ? hitHandle(p) : -1;
}

LPCWSTR cursorOf(const OverlayHandle& h) {
    switch (h.action) {
        case OverlayAction::RowMove:
        case OverlayAction::GanttMove: return IDC_SIZEALL;
        case OverlayAction::GanttStart:
        case OverlayAction::GanttEnd:
        case OverlayAction::RangeStart:
        case OverlayAction::RangeEnd: return IDC_SIZEWE;
        case OverlayAction::CellEdit: return IDC_IBEAM;
        case OverlayAction::Datasheet: return IDC_HAND;
        default: return h.dirX != 0 ? IDC_SIZEWE : IDC_SIZENS;
    }
}

void post(HWND hwnd, OverlayGesture gesture, int pixels) {
    PostMessageW(hwnd, kMsgAction, static_cast<WPARAM>(gesture), static_cast<LPARAM>(pixels));
}

void render(bool move);

// Maus über einem anderen Griff (oder keinem): neu zeichnen
void updateHover() {
    const int hover = hoverTarget();
    if (hover == g_hover) return;
    g_hover = hover;
    if (g_window && IsWindowVisible(g_window)) render(false);
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
            if (g_press.active) {
                SetCursor(LoadCursorW(nullptr, cursorOf(g_press.handle)));
            } else {
                const int index = hitHandle(messagePos());
                SetCursor(LoadCursorW(nullptr, index < 0 ? IDC_HAND : cursorOf(g_content.handles[index])));
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
            if (!g_press.active) {
                updateHover();
                TRACKMOUSEEVENT track{sizeof(track), TME_LEAVE, hwnd, 0};
                TrackMouseEvent(&track);
                return 0;
            }
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
            SetCursor(LoadCursorW(nullptr, cursorOf(h)));
            if (!g_press.posted) {
                g_press.posted = true;
                post(hwnd, OverlayGesture::Drag, pixels);
            }
            return 0;
        }
        case WM_MOUSELEAVE:
            if (!g_press.active) updateHover();
            return 0;
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
            if (!g_press.active) updateHover();  // nach dem Loslassen: die Maus liegt evtl. woanders
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

// Sechs Punkte „⋮⋮“ (zwei Spalten, drei Reihen) mittig im Rechteck
void drawDots(Canvas& canvas, const RECT& r, int dot, double alpha) {
    const int gap = dot, w = 2 * dot + gap, h = 3 * dot + 2 * gap;
    const int x0 = (r.left + r.right - w) / 2, y0 = (r.top + r.bottom - h) / 2;
    for (int i = 0; i < 2; ++i)
        for (int j = 0; j < 3; ++j) {
            const int x = x0 + i * (dot + gap), y = y0 + j * (dot + gap);
            canvas.fill({x, y, x + dot, y + dot}, kBlue, alpha);
        }
}

// Rechteck mit 1 Pixel breitem Rand (Fensterkoordinaten)
void frame(Canvas& canvas, const RECT& r, double alpha) {
    canvas.fill({r.left, r.top, r.right, r.top + 1}, kBlue, alpha);
    canvas.fill({r.left, r.bottom - 1, r.right, r.bottom}, kBlue, alpha);
    canvas.fill({r.left, r.top, r.left + 1, r.bottom}, kBlue, alpha);
    canvas.fill({r.right - 1, r.top, r.right, r.bottom}, kBlue, alpha);
}

// Zeichnet g_content in das Ebenenfenster. move = Fenster an g_origin setzen (sonst Lage belassen, z. B. beim Überfahren).
void render(bool move) {
    HWND hwnd = g_window;
    if (!hwnd) return;
    HWND owner = GetWindow(hwnd, GW_OWNER);
    const OverlayContent& content = g_content;

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

    // Skalenband: halbtransparent mit kräftigen Kanten oben und unten
    if (!IsRectEmpty(&content.band)) {
        RECT r = content.band;
        OffsetRect(&r, -g_origin.x, -g_origin.y);
        canvas.fill(r, kBlue, 0.16);
        canvas.fill({r.left, r.top, r.right, r.top + 1}, kBlue, 0.9);
        canvas.fill({r.left, r.bottom - 1, r.right, r.bottom}, kBlue, 0.9);
    }

    // Maus über einer Zeile (Griff oder Zelle): sehr heller Streifen über ihrem Beschriftungsbereich, eine Zelle leuchtet mit Rahmen
    const OverlayHandle* hot = g_hover >= 0 && g_hover < static_cast<int>(content.handles.size()) ? &content.handles[g_hover] : nullptr;
    const bool hotRow = hot && (hot->action == OverlayAction::RowMove || hot->action == OverlayAction::CellEdit);
    if (hotRow) {
        RECT area = hot->area;
        OffsetRect(&area, -g_origin.x, -g_origin.y);
        canvas.fill(area, kBlue, 0.07);
        if (hot->action == OverlayAction::CellEdit) {
            RECT r = IsRectEmpty(&hot->cell) ? hot->rect : hot->cell;
            OffsetRect(&r, -g_origin.x, -g_origin.y);
            canvas.fill(r, kBlue, 0.16);
            frame(canvas, r, 0.9);
        }
    }

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
            case HandleStyle::DoubleArrow: {
                const int mid = (r.left + r.right) / 2;
                canvas.triangle({r.left, r.top, mid, r.bottom}, 2);
                canvas.triangle({mid, r.top, r.right, r.bottom}, 3);
                break;
            }
            case HandleStyle::Grip: canvas.grip(r); break;
            case HandleStyle::Hidden: canvas.fill(r, kBlue, 0.004); break;  // fast unsichtbar, aber für das Fenster vorhanden
            case HandleStyle::RowGrip: {
                const bool lit = hotRow && hot->row == h.row;
                canvas.fill(r, kBlue, lit ? 0.14 : 0.004);  // auch unscheinbar „da“, damit das Fenster die Maus bekommt
                drawDots(canvas, r, std::max(2, scaled(owner, 2)), lit ? 1.0 : 0.4);
                break;
            }
        }
    }

    // Grenze unter der Maus (Zeilen-/Skalenrand): helles Band, an dem eine Trennlinie eingefügt werden kann
    if (!IsRectEmpty(&content.line)) {
        RECT r = content.line;
        OffsetRect(&r, -g_origin.x, -g_origin.y);
        canvas.fill(r, kBlue, 0.3);
    }

    // Einfügelinie beim Ziehen einer Zeile
    if (!IsRectEmpty(&content.marker)) {
        RECT r = content.marker;
        OffsetRect(&r, -g_origin.x, -g_origin.y);
        canvas.fill(r, kBlue, 1.0);
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
    UpdateLayeredWindow(hwnd, screen, move ? &position : nullptr, &size, dc, &source, 0, &blend, ULW_ALPHA);
    if (move) ShowWindow(hwnd, SW_SHOWNOACTIVATE);

    SelectObject(dc, previous);
    DeleteObject(bitmap);
    DeleteDC(dc);
    ReleaseDC(nullptr, screen);
}

}  // namespace

void showOverlay(HWND owner, const OverlayContent& content, OverlayCallback callback) {
    HWND hwnd = ensureWindow(owner);
    if (!hwnd) return;
    g_content = content;
    g_callback = callback;
    g_displaced = false;
    g_hover = -1;
    render(true);
    g_hover = hoverTarget();  // die Maus liegt nach dem Neuaufbau evtl. schon über einem Griff
    if (g_hover >= 0) render(false);
}

void overlayHoverCheck() {
    if (g_window && IsWindowVisible(g_window) && !g_press.active) updateHover();
}

void setOverlayLine(const RECT& line) {
    const bool empty = IsRectEmpty(&line) != 0;
    if (empty ? IsRectEmpty(&g_content.line) != 0 : EqualRect(&g_content.line, &line) != 0) return;
    g_content.line = empty ? RECT{} : line;
    if (g_window && IsWindowVisible(g_window)) render(false);
}

const OverlayHandle* overlayHoverAt(POINT screen) {
    g_hover = hitHandle(screen);
    if (g_window && IsWindowVisible(g_window)) render(false);
    return g_hover >= 0 ? &g_content.handles[g_hover] : nullptr;
}

void hideOverlay() {
    if (g_window) {
        if (g_press.active) release(g_window, OverlayGesture::DragCancel);
        ShowWindow(g_window, SW_HIDE);
    }
    g_content = {};
    g_hover = -1;
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
