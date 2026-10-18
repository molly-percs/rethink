#include "calendar_window.h"

#include "engine/dates.h"

#include <commctrl.h>

#include <algorithm>
#include <cmath>
#include <format>

extern "C" IMAGE_DOS_HEADER __ImageBase;

namespace rethink::addin {
namespace {

constexpr wchar_t kClass[] = L"rethinkCalendar";
constexpr int kIdCalendar = 10, kIdClose = 11, kIdRange = 12;

struct State {
    HWND window = nullptr, calendar = nullptr, close = nullptr, range = nullptr;
    HFONT font = nullptr;
    bool single = false;
    CalendarCallbacks callbacks;
};
State g_state;

SYSTEMTIME systemTime(double t) {
    const engine::CivilDate c = engine::civilFromDays(static_cast<long long>(std::floor(t)));
    SYSTEMTIME st{};
    st.wYear = static_cast<WORD>(c.year), st.wMonth = static_cast<WORD>(c.month), st.wDay = static_cast<WORD>(c.day);
    st.wDayOfWeek = static_cast<WORD>(engine::isoWeekday(static_cast<long long>(std::floor(t))) % 7);
    return st;
}

double daysOf(const SYSTEMTIME& st) { return static_cast<double>(engine::daysFromCivil(st.wYear, st.wMonth, st.wDay)); }

void showRange(double first, double last) {
    const long long days = std::llround(last - first) + 1;
    const std::wstring text = g_state.single ? engine::isoDate(first)
                                             : std::format(L"{} – {}  ({} {})", engine::isoDate(first), engine::isoDate(last), days, days == 1 ? L"Tag" : L"Tage");
    SetWindowTextW(g_state.range, text.c_str());
}

void selected() {
    SYSTEMTIME st[2]{};
    double first = 0, last = 0;
    if (g_state.single) {
        MonthCal_GetCurSel(g_state.calendar, &st[0]);
        first = last = daysOf(st[0]);
    } else {
        MonthCal_GetSelRange(g_state.calendar, st);
        first = daysOf(st[0]), last = daysOf(st[1]);
    }
    showRange(first, last);
    auto changed = g_state.callbacks.changed;  // der Rückruf zeichnet in PowerPoint neu
    if (changed) changed(first, last);
    if (g_state.single) PostMessageW(g_state.window, WM_CLOSE, 0, 0);
}

LRESULT CALLBACK windowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_NOTIFY: {
            const auto* header = reinterpret_cast<const NMHDR*>(lParam);
            if (header->idFrom == kIdCalendar && header->code == MCN_SELECT) {
                try {
                    selected();
                } catch (...) {
                }
            }
            return 0;
        }
        case WM_COMMAND:
            if (LOWORD(wParam) == kIdClose || LOWORD(wParam) == IDCANCEL) DestroyWindow(hwnd);
            return 0;
        case WM_ACTIVATE:
            // Klick in PowerPoint schließt den Kalender (wie im Original)
            if (LOWORD(wParam) == WA_INACTIVE && lParam && GetWindow(hwnd, GW_OWNER) == reinterpret_cast<HWND>(lParam)) PostMessageW(hwnd, WM_CLOSE, 0, 0);
            return 0;
        case WM_CLOSE: DestroyWindow(hwnd); return 0;
        case WM_NCDESTROY:
            if (g_state.font) DeleteObject(g_state.font);
            g_state = {};
            return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

}  // namespace

void openCalendarWindow(HWND owner, const std::wstring& title, double first, double last, bool single, CalendarCallbacks callbacks,
                        const RECT& beside) {
    closeCalendarWindow();
    INITCOMMONCONTROLSEX icc{sizeof(icc), ICC_DATE_CLASSES | ICC_STANDARD_CLASSES};
    InitCommonControlsEx(&icc);
    const auto module = reinterpret_cast<HINSTANCE>(&__ImageBase);
    static bool registered = false;
    if (!registered) {
        WNDCLASSEXW wc{sizeof(wc)};
        wc.lpfnWndProc = windowProc;
        wc.hInstance = module;
        wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
        wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
        wc.lpszClassName = kClass;
        RegisterClassExW(&wc);
        registered = true;
    }
    HWND hwnd = CreateWindowExW(WS_EX_TOOLWINDOW | WS_EX_DLGMODALFRAME, kClass, title.c_str(), WS_POPUP | WS_CAPTION | WS_SYSMENU, 0, 0, 100, 100,
                                owner, nullptr, module, nullptr);
    if (!hwnd) return;
    g_state.window = hwnd;
    g_state.single = single;
    g_state.callbacks = std::move(callbacks);
    const UINT dpi = GetDpiForWindow(hwnd) ? GetDpiForWindow(hwnd) : 96;
    g_state.font = CreateFontW(-MulDiv(9, dpi, 72), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                               CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");

    const int pad = MulDiv(8, dpi, 96), lineH = MulDiv(22, dpi, 96), buttonW = MulDiv(90, dpi, 96), buttonH = MulDiv(26, dpi, 96);
    g_state.calendar = CreateWindowExW(0, MONTHCAL_CLASSW, L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | MCS_WEEKNUMBERS | (single ? 0 : MCS_MULTISELECT),
                                       pad, pad + lineH, 10, 10, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kIdCalendar)), module, nullptr);
    SendMessageW(g_state.calendar, WM_SETFONT, reinterpret_cast<WPARAM>(g_state.font), TRUE);
    MonthCal_SetFirstDayOfWeek(g_state.calendar, 0);  // Montag
    if (!single) MonthCal_SetMaxSelCount(g_state.calendar, 3660);
    RECT one{};
    MonthCal_GetMinReqRect(g_state.calendar, &one);
    // drei Monate nebeneinander, zwei Reihen (Größe mit der Breite des Fensters wählbar wäre ein nächster Schritt)
    RECT want{0, 0, (one.right - one.left) * 3, (one.bottom - one.top) * 2};
    MonthCal_SizeRectToMin(g_state.calendar, &want);
    const int calW = want.right - want.left, calH = want.bottom - want.top;
    SetWindowPos(g_state.calendar, nullptr, pad, pad + lineH, calW, calH, SWP_NOZORDER);

    g_state.range = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_RIGHT, pad, pad, calW, lineH, hwnd,
                                    reinterpret_cast<HMENU>(static_cast<INT_PTR>(kIdRange)), module, nullptr);
    SendMessageW(g_state.range, WM_SETFONT, reinterpret_cast<WPARAM>(g_state.font), TRUE);
    g_state.close = CreateWindowExW(0, L"BUTTON", L"Schließen", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, pad + calW - buttonW,
                                    pad + lineH + calH + pad, buttonW, buttonH, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kIdClose)), module,
                                    nullptr);
    SendMessageW(g_state.close, WM_SETFONT, reinterpret_cast<WPARAM>(g_state.font), TRUE);

    const SYSTEMTIME a = systemTime(first), b = systemTime(std::max(first, last));
    if (single) {
        MonthCal_SetCurSel(g_state.calendar, &a);
    } else {
        const SYSTEMTIME range[2] = {a, b};
        MonthCal_SetSelRange(g_state.calendar, range);
    }
    showRange(first, std::max(first, last));

    RECT frame{0, 0, calW + 2 * pad, pad + lineH + calH + pad + buttonH + pad};
    AdjustWindowRectExForDpi(&frame, WS_POPUP | WS_CAPTION | WS_SYSMENU, FALSE, WS_EX_TOOLWINDOW | WS_EX_DLGMODALFRAME, dpi);
    const int w = frame.right - frame.left, h = frame.bottom - frame.top;
    MONITORINFO mi{sizeof(mi)};
    GetMonitorInfoW(MonitorFromRect(&beside, MONITOR_DEFAULTTONEAREST), &mi);
    int x = std::clamp<int>(beside.left, mi.rcWork.left, mi.rcWork.right - w);
    int y = beside.bottom + 8;
    if (y + h > mi.rcWork.bottom) y = std::max<int>(mi.rcWork.top, beside.top - h - 8);
    SetWindowPos(hwnd, HWND_TOP, x, y, w, h, SWP_SHOWWINDOW);
    SetFocus(g_state.calendar);
}

void closeCalendarWindow() {
    if (g_state.window) DestroyWindow(g_state.window);
}

bool calendarWindowOpen() { return g_state.window != nullptr; }

}  // namespace rethink::addin
