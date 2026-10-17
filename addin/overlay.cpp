#include "overlay.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <vector>

extern "C" IMAGE_DOS_HEADER __ImageBase;

namespace rethink::addin {
namespace {

constexpr wchar_t kClass[] = L"rethinkSelectionOverlay";
constexpr int kGlow = 10;                       // Leuchtbreite außerhalb des Elements in Pixeln
constexpr COLORREF kBlue = RGB(37, 120, 235);   // rethink-Blau
constexpr int kBadgeHeight = 18;

HWND g_window = nullptr;

LRESULT CALLBACK windowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    if (msg == WM_NCHITTEST) return HTTRANSPARENT;
    if (msg == WM_MOUSEACTIVATE) return MA_NOACTIVATE;
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
    g_window = CreateWindowExW(WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW, kClass, L"", WS_POPUP, 0, 0, 1, 1,
                               owner, nullptr, module, nullptr);
    return g_window;
}

// Vorbelegtes (premultiplied) BGRA-Pixel
std::uint32_t pixel(COLORREF color, double alpha) {
    const auto a = static_cast<std::uint32_t>(std::clamp(alpha, 0.0, 1.0) * 255.0 + 0.5);
    const auto r = GetRValue(color) * a / 255, g = GetGValue(color) * a / 255, b = GetBValue(color) * a / 255;
    return (a << 24) | (r << 16) | (g << 8) | b;
}

}  // namespace

void showOverlay(HWND owner, const RECT& rect, const std::wstring& label) {
    HWND hwnd = ensureWindow(owner);
    if (!hwnd) return;

    // Fenster = Element + Leuchtrand + Platz für das Etikett oberhalb
    const int innerW = std::max<int>(rect.right - rect.left, 1), innerH = std::max<int>(rect.bottom - rect.top, 1);
    const int width = innerW + 2 * kGlow, height = innerH + 2 * kGlow + kBadgeHeight;
    const int originX = rect.left - kGlow, originY = rect.top - kGlow - kBadgeHeight;

    BITMAPINFO info{};
    info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    info.bmiHeader.biWidth = width;
    info.bmiHeader.biHeight = -height;  // von oben nach unten
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

    // Leuchtrahmen: kräftige Linie am Elementrand, nach außen weich auslaufend, innen dezenter Schimmer
    const int left = kGlow, top = kGlow + kBadgeHeight, right = left + innerW, bottom = top + innerH;
    for (int y = kBadgeHeight; y < height; ++y) {
        for (int x = 0; x < width; ++x) {
            const int dx = x < left ? left - x : (x >= right ? x - right + 1 : 0);
            const int dy = y < top ? top - y : (y >= bottom ? y - bottom + 1 : 0);
            double alpha = 0;
            if (dx == 0 && dy == 0) {
                const int inside = std::min({x - left, right - 1 - x, y - top, bottom - 1 - y});
                if (inside < 2) alpha = 0.95;                 // Kante
                else if (inside < 6) alpha = 0.18 * (6 - inside) / 4.0;  // Schimmer nach innen
            } else {
                const double d = std::sqrt(static_cast<double>(dx * dx + dy * dy));
                if (d <= kGlow) alpha = 0.55 * std::pow(1.0 - d / kGlow, 1.6);
            }
            px[y * width + x] = pixel(kBlue, alpha);
        }
    }

    // Etikett oben links: blaue Fläche mit weißem Text
    HFONT font = CreateFontW(-12, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                             CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    HGDIOBJ oldFont = SelectObject(dc, font);
    SIZE textSize{};
    GetTextExtentPoint32W(dc, label.c_str(), static_cast<int>(label.size()), &textSize);
    RECT badge{left, 0, std::min(left + static_cast<int>(textSize.cx) + 14, width), kBadgeHeight - 2};
    HBRUSH brush = CreateSolidBrush(kBlue);
    FillRect(dc, &badge, brush);
    DeleteObject(brush);
    SetBkMode(dc, TRANSPARENT);
    SetTextColor(dc, RGB(255, 255, 255));
    RECT textRect = badge;
    DrawTextW(dc, label.c_str(), static_cast<int>(label.size()), &textRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
    SelectObject(dc, oldFont);
    DeleteObject(font);
    for (int y = badge.top; y < badge.bottom; ++y)
        for (int x = badge.left; x < badge.right; ++x) px[y * width + x] |= 0xFF000000;  // GDI setzt kein Alpha

    POINT position{originX, originY}, source{0, 0};
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
    if (g_window) ShowWindow(g_window, SW_HIDE);
}

bool overlayVisible() { return g_window && IsWindowVisible(g_window); }

}  // namespace rethink::addin
