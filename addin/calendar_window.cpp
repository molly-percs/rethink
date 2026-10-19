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
constexpr int kIdClose = 11, kIdPrev = 12, kIdNext = 13, kIdMonth = 14, kIdYear = 15, kIdSpin = 16;
constexpr int kMonths = 12, kColumns = 37;  // ein Jahr ab dem gewählten Monat; 37 Spalten reichen für jeden Monat ab Montag
constexpr COLORREF kBlue = RGB(37, 120, 235), kLightBlue = RGB(204, 224, 250), kWeekend = RGB(240, 240, 240), kRed = RGB(210, 20, 20);
const wchar_t* const kMonthNames[] = {L"Januar", L"Februar", L"März",     L"April",   L"Mai",      L"Juni",
                                      L"Juli",   L"August",  L"September", L"Oktober", L"November", L"Dezember"};
const wchar_t* const kWeekdays[] = {L"Mo", L"Di", L"Mi", L"Do", L"Fr", L"Sa", L"So"};

struct State {
    HWND window = nullptr, close = nullptr, month = nullptr, year = nullptr, spin = nullptr;
    HFONT font = nullptr, bold = nullptr;
    bool single = false;
    CalendarCallbacks callbacks;
    int startYear = 2026, startMonth = 1;  // oberste Zeile
    double first = 0, last = 0, today = 0;
    int hover = -1;  // Tag unter der Maus (Tage seit 1970), -1 = keiner
    enum class Drag { None, New, First, Last } drag = Drag::None;
    double anchor = 0;
    bool updating = false;  // Steuerelemente werden gerade gesetzt: deren Meldungen nicht auswerten
    bool applying = false;  // der Rückruf zeichnet in PowerPoint neu und aktiviert es dabei: nicht schließen
    int dpi = 96;
    int s(int v) const { return MulDiv(v, dpi, 96); }
    // Raster in Client-Koordinaten
    int gridLeft() const { return s(10) + s(78); }
    int gridTop() const { return s(10) + s(30) + s(18); }
    int cellW() const { return s(22); }
    int cellH() const { return s(20); }
    int gridRight() const { return gridLeft() + kColumns * cellW(); }
    int gridBottom() const { return gridTop() + kMonths * cellH(); }
};
State g_state;

int daysInMonth(int year, int month) {
    const int nextYear = month == 12 ? year + 1 : year, next = month == 12 ? 1 : month + 1;
    return static_cast<int>(engine::daysFromCivil(nextYear, next, 1) - engine::daysFromCivil(year, month, 1));
}

void monthOfRow(int row, int& year, int& month) {
    const int index = g_state.startMonth - 1 + row;
    year = g_state.startYear + index / 12;
    month = index % 12 + 1;
}

// Tag an einem Client-Punkt (Tage seit 1970), -1 = kein Tag
long long dayAt(POINT p) {
    if (p.x < g_state.gridLeft() || p.x >= g_state.gridRight() || p.y < g_state.gridTop() || p.y >= g_state.gridBottom()) return -1;
    const int row = (p.y - g_state.gridTop()) / g_state.cellH(), column = (p.x - g_state.gridLeft()) / g_state.cellW();
    int year = 0, month = 0;
    monthOfRow(row, year, month);
    const long long firstDay = engine::daysFromCivil(year, month, 1);
    const int day = column - (engine::isoWeekday(firstDay) - 1) + 1;
    return day >= 1 && day <= daysInMonth(year, month) ? firstDay + day - 1 : -1;
}

std::wstring dayText(double t) {
    const auto day = static_cast<long long>(std::floor(t));
    return std::format(L"{} {} · KW {}", kWeekdays[engine::isoWeekday(day) - 1], engine::formatDate(t, L"dd.MM.yyyy", L"de"), engine::isoWeek(day));
}

// „01.06.2026 bis 30.06.2026 · 4 Wochen 2 Tage“ bzw. beim Meilenstein der Tag mit Kalenderwoche
std::wstring summaryText() {
    if (g_state.single) return dayText(g_state.first);
    const long long days = std::llround(g_state.last - g_state.first) + 1, weeks = days / 7, rest = days % 7;
    std::wstring length;
    if (weeks > 0) length = std::format(L"{} {}", weeks, weeks == 1 ? L"Woche" : L"Wochen");
    if (rest > 0 || weeks == 0) length += std::format(L"{}{} {}", length.empty() ? L"" : L" ", rest, rest == 1 ? L"Tag" : L"Tage");
    return std::format(L"{} bis {} · {}", engine::formatDate(g_state.first, L"dd.MM.yyyy", L"de"), engine::formatDate(g_state.last, L"dd.MM.yyyy", L"de"),
                       length);
}

void syncControls() {
    g_state.updating = true;
    SendMessageW(g_state.month, CB_SETCURSEL, g_state.startMonth - 1, 0);
    SendMessageW(g_state.spin, UDM_SETPOS32, 0, g_state.startYear);
    g_state.updating = false;
}

void scrollTo(int year, int month) {
    g_state.startYear = std::clamp(year, 1901, 2199), g_state.startMonth = month;
    syncControls();
    InvalidateRect(g_state.window, nullptr, FALSE);
}

void shiftMonths(int n) {
    const int index = g_state.startYear * 12 + g_state.startMonth - 1 + n;
    scrollTo(index / 12, index % 12 + 1);
}

void apply() {
    InvalidateRect(g_state.window, nullptr, FALSE);
    auto changed = g_state.callbacks.changed;  // der Rückruf zeichnet in PowerPoint neu
    HWND hwnd = g_state.window;
    g_state.applying = true;
    if (changed) changed(g_state.first, g_state.last);
    if (!IsWindow(hwnd)) return;
    g_state.applying = false;
    if (g_state.single) PostMessageW(hwnd, WM_CLOSE, 0, 0);
    else SetActiveWindow(hwnd);
}

void fillRect(HDC dc, RECT r, COLORREF c) {
    HBRUSH b = CreateSolidBrush(c);
    FillRect(dc, &r, b);
    DeleteObject(b);
}

void paint(HWND hwnd) {
    PAINTSTRUCT ps;
    HDC target = BeginPaint(hwnd, &ps);
    RECT client;
    GetClientRect(hwnd, &client);
    HDC dc = CreateCompatibleDC(target);
    HBITMAP bitmap = CreateCompatibleBitmap(target, client.right, client.bottom);
    HGDIOBJ oldBitmap = SelectObject(dc, bitmap);
    fillRect(dc, client, GetSysColor(COLOR_WINDOW));
    SetBkMode(dc, TRANSPARENT);
    HGDIOBJ oldFont = SelectObject(dc, g_state.font);
    const State& s = g_state;
    const int cw = s.cellW(), ch = s.cellH();

    // Kopf: Wochentage, Wochenenden grau
    const int headTop = s.gridTop() - s.s(18);
    for (int c = 0; c < kColumns; ++c) {
        RECT r{s.gridLeft() + c * cw, headTop, s.gridLeft() + (c + 1) * cw, s.gridBottom()};
        if (c % 7 >= 5) fillRect(dc, r, kWeekend);
        r.bottom = s.gridTop();
        SetTextColor(dc, RGB(110, 110, 110));
        DrawTextW(dc, kWeekdays[c % 7], -1, &r, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    }
    // Monatszeilen
    for (int row = 0; row < kMonths; ++row) {
        int year = 0, month = 0;
        monthOfRow(row, year, month);
        const int top = s.gridTop() + row * ch;
        RECT label{s.s(10), top, s.gridLeft() - s.s(6), top + ch};
        SetTextColor(dc, GetSysColor(COLOR_WINDOWTEXT));
        SelectObject(dc, month == 1 || row == 0 ? s.bold : s.font);
        const std::wstring name = std::format(L"{} {}", std::wstring(kMonthNames[month - 1]).substr(0, 3), year);
        DrawTextW(dc, name.c_str(), -1, &label, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
        SelectObject(dc, s.font);
        const long long firstDay = engine::daysFromCivil(year, month, 1);
        const int offset = engine::isoWeekday(firstDay) - 1, length = daysInMonth(year, month);
        for (int d = 1; d <= length; ++d) {
            const long long day = firstDay + d - 1;
            RECT r{s.gridLeft() + (offset + d - 1) * cw, top, s.gridLeft() + (offset + d) * cw, top + ch};
            const bool selected = day >= s.first && day <= s.last, end = day == s.first || day == s.last;
            RECT inner{r.left + 1, r.top + 1, r.right - 1, r.bottom - 1};
            if (selected) fillRect(dc, inner, end ? kBlue : kLightBlue);
            if (day == s.hover && !selected) {
                HBRUSH frame = CreateSolidBrush(kBlue);
                FrameRect(dc, &inner, frame);
                DeleteObject(frame);
            }
            const bool today = day == static_cast<long long>(s.today);
            SetTextColor(dc, selected && end ? RGB(255, 255, 255) : today ? kRed : GetSysColor(COLOR_WINDOWTEXT));
            SelectObject(dc, today ? s.bold : s.font);
            DrawTextW(dc, std::to_wstring(d).c_str(), -1, &r, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            if (today) {
                HBRUSH red = CreateSolidBrush(kRed);
                RECT mark{r.left + s.s(4), r.bottom - s.s(3), r.right - s.s(4), r.bottom - s.s(2)};
                FillRect(dc, &mark, red);
                DeleteObject(red);
            }
        }
        SelectObject(dc, s.font);
        // feine Linie zwischen den Monaten
        fillRect(dc, {s.gridLeft(), top + ch - 1, s.gridRight(), top + ch}, RGB(228, 228, 228));
    }
    // Zusammenfassung und Tag unter der Maus
    const int statusTop = s.gridBottom() + s.s(8);
    RECT summary{s.s(10), statusTop, s.gridRight(), statusTop + s.s(20)};
    SetTextColor(dc, GetSysColor(COLOR_WINDOWTEXT));
    SelectObject(dc, s.bold);
    DrawTextW(dc, summaryText().c_str(), -1, &summary, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
    SelectObject(dc, s.font);
    if (s.hover >= 0) {
        SetTextColor(dc, RGB(90, 90, 90));
        DrawTextW(dc, dayText(s.hover).c_str(), -1, &summary, DT_RIGHT | DT_VCENTER | DT_SINGLELINE);
    }
    SelectObject(dc, oldFont);
    BitBlt(target, 0, 0, client.right, client.bottom, dc, 0, 0, SRCCOPY);
    SelectObject(dc, oldBitmap);
    DeleteObject(bitmap);
    DeleteDC(dc);
    EndPaint(hwnd, &ps);
}

void mouseDown(POINT p, bool shift) {
    const long long day = dayAt(p);
    if (day < 0) return;
    SetCapture(g_state.window);
    const double d = static_cast<double>(day);
    if (g_state.single) {
        g_state.drag = State::Drag::New;
        g_state.first = g_state.last = d;
    } else if (shift) {  // Umschalt: den näheren Rand bis hierher ziehen
        g_state.drag = std::fabs(d - g_state.first) < std::fabs(d - g_state.last) ? State::Drag::First : State::Drag::Last;
        (g_state.drag == State::Drag::First ? g_state.first : g_state.last) = d;
    } else if (d == g_state.first && g_state.first != g_state.last) {
        g_state.drag = State::Drag::First;  // Anfang des Bereichs greifen
    } else if (d == g_state.last && g_state.first != g_state.last) {
        g_state.drag = State::Drag::Last;
    } else {
        g_state.drag = State::Drag::New;
        g_state.anchor = g_state.first = g_state.last = d;
    }
    InvalidateRect(g_state.window, nullptr, FALSE);
}

void mouseMove(POINT p) {
    const long long day = dayAt(p);
    if (day != g_state.hover) {
        g_state.hover = static_cast<int>(day);
        InvalidateRect(g_state.window, nullptr, FALSE);
    }
    if (g_state.drag == State::Drag::None || day < 0) return;
    const double d = static_cast<double>(day);
    switch (g_state.drag) {
        case State::Drag::New:
            if (g_state.single) g_state.first = g_state.last = d;
            else g_state.first = std::min(g_state.anchor, d), g_state.last = std::max(g_state.anchor, d);
            break;
        case State::Drag::First: g_state.first = std::min(d, g_state.last); break;
        case State::Drag::Last: g_state.last = std::max(d, g_state.first); break;
        case State::Drag::None: break;
    }
}

void mouseUp() {
    if (g_state.drag == State::Drag::None) return;
    g_state.drag = State::Drag::None;
    ReleaseCapture();
    try {
        apply();
    } catch (...) {
    }
}

LRESULT CALLBACK windowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    const POINT p{static_cast<short>(LOWORD(lParam)), static_cast<short>(HIWORD(lParam))};
    switch (msg) {
        case WM_PAINT: paint(hwnd); return 0;
        case WM_ERASEBKGND: return 1;
        case WM_LBUTTONDOWN: mouseDown(p, (wParam & MK_SHIFT) != 0); return 0;
        case WM_MOUSEMOVE: {
            mouseMove(p);
            TRACKMOUSEEVENT track{sizeof(track), TME_LEAVE, hwnd, 0};
            TrackMouseEvent(&track);
            return 0;
        }
        case WM_MOUSELEAVE:
            if (g_state.hover >= 0) g_state.hover = -1, InvalidateRect(hwnd, nullptr, FALSE);
            return 0;
        case WM_LBUTTONUP: mouseUp(); return 0;
        case WM_MOUSEWHEEL: shiftMonths(GET_WHEEL_DELTA_WPARAM(wParam) > 0 ? -1 : 1); return 0;
        case WM_KEYDOWN:
            if (wParam == VK_ESCAPE) DestroyWindow(hwnd);
            return 0;
        case WM_COMMAND:
            switch (LOWORD(wParam)) {
                case kIdClose:
                case IDCANCEL: DestroyWindow(hwnd); break;
                case kIdPrev: shiftMonths(-1); break;
                case kIdNext: shiftMonths(1); break;
                case kIdMonth:
                    if (HIWORD(wParam) == CBN_SELCHANGE && !g_state.updating)
                        scrollTo(g_state.startYear, static_cast<int>(SendMessageW(g_state.month, CB_GETCURSEL, 0, 0)) + 1);
                    break;
                case kIdYear:
                    if (HIWORD(wParam) == EN_CHANGE && !g_state.updating) {
                        BOOL failed = FALSE;
                        const int year = static_cast<int>(SendMessageW(g_state.spin, UDM_GETPOS32, 0, reinterpret_cast<LPARAM>(&failed)));
                        if (!failed && year != g_state.startYear) scrollTo(year, g_state.startMonth);
                    }
                    break;
            }
            return 0;
        case WM_ACTIVATE:
            // Klick in PowerPoint schließt den Kalender (wie im Original)
            if (LOWORD(wParam) == WA_INACTIVE && !g_state.applying && lParam && GetWindow(hwnd, GW_OWNER) == reinterpret_cast<HWND>(lParam))
                PostMessageW(hwnd, WM_CLOSE, 0, 0);
            return 0;
        case WM_CLOSE: DestroyWindow(hwnd); return 0;
        case WM_NCDESTROY:
            if (g_state.font) DeleteObject(g_state.font);
            if (g_state.bold) DeleteObject(g_state.bold);
            g_state = {};
            return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

HWND child(const wchar_t* cls, const wchar_t* text, DWORD style, int id, int x, int y, int w, int h) {
    HWND c = CreateWindowExW(0, cls, text, WS_CHILD | WS_VISIBLE | style, x, y, w, h, g_state.window, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),
                             reinterpret_cast<HINSTANCE>(&__ImageBase), nullptr);
    SendMessageW(c, WM_SETFONT, reinterpret_cast<WPARAM>(g_state.font), TRUE);
    return c;
}

}  // namespace

void openCalendarWindow(HWND owner, const std::wstring& title, double first, double last, bool single, CalendarCallbacks callbacks,
                        const RECT& beside) {
    closeCalendarWindow();
    INITCOMMONCONTROLSEX icc{sizeof(icc), ICC_UPDOWN_CLASS | ICC_STANDARD_CLASSES};
    InitCommonControlsEx(&icc);
    const auto module = reinterpret_cast<HINSTANCE>(&__ImageBase);
    static bool registered = false;
    if (!registered) {
        WNDCLASSEXW wc{sizeof(wc)};
        wc.lpfnWndProc = windowProc;
        wc.hInstance = module;
        wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kClass;
        RegisterClassExW(&wc);
        registered = true;
    }
    HWND hwnd = CreateWindowExW(WS_EX_TOOLWINDOW | WS_EX_DLGMODALFRAME, kClass, title.c_str(), WS_POPUP | WS_CAPTION | WS_SYSMENU, 0, 0, 100, 100,
                                owner, nullptr, module, nullptr);
    if (!hwnd) return;
    State& s = g_state;
    s.window = hwnd;
    s.single = single;
    s.callbacks = std::move(callbacks);
    s.dpi = GetDpiForWindow(hwnd) ? static_cast<int>(GetDpiForWindow(hwnd)) : 96;
    s.font = CreateFontW(-MulDiv(9, s.dpi, 72), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                         CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    s.bold = CreateFontW(-MulDiv(9, s.dpi, 72), 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                         CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    s.first = std::floor(first), s.last = std::max(s.first, std::floor(last));
    SYSTEMTIME now{};
    GetLocalTime(&now);
    s.today = static_cast<double>(engine::daysFromCivil(now.wYear, now.wMonth, now.wDay));
    const engine::CivilDate start = engine::civilFromDays(static_cast<long long>(s.first));
    s.startYear = start.year, s.startMonth = start.month;

    // Kopfzeile: Monat zurück/vor, Monat und Jahr der obersten Zeile
    const int pad = s.s(10), barH = s.s(24);
    child(L"BUTTON", L"◀", BS_PUSHBUTTON, kIdPrev, pad, pad, s.s(26), barH);
    child(L"BUTTON", L"▶", BS_PUSHBUTTON, kIdNext, pad + s.s(28), pad, s.s(26), barH);
    s.month = child(L"COMBOBOX", L"", CBS_DROPDOWNLIST | WS_VSCROLL | WS_TABSTOP, kIdMonth, pad + s.s(62), pad, s.s(110), s.s(300));
    for (const wchar_t* name : kMonthNames) SendMessageW(s.month, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(name));
    s.year = child(L"EDIT", L"", ES_NUMBER | WS_BORDER | WS_TABSTOP, kIdYear, pad + s.s(180), pad, s.s(56), barH);
    s.spin = child(UPDOWN_CLASSW, L"", UDS_SETBUDDYINT | UDS_ALIGNRIGHT | UDS_ARROWKEYS | UDS_NOTHOUSANDS, kIdSpin, 0, 0, 0, 0);
    SendMessageW(s.spin, UDM_SETBUDDY, reinterpret_cast<WPARAM>(s.year), 0);
    SendMessageW(s.spin, UDM_SETRANGE32, 1901, 2199);
    syncControls();

    const int buttonW = s.s(90), buttonH = s.s(26), statusBottom = s.gridBottom() + s.s(8) + s.s(20);
    s.close = child(L"BUTTON", L"Schließen", BS_PUSHBUTTON | WS_TABSTOP, kIdClose, s.gridRight() - buttonW, statusBottom + s.s(8), buttonW, buttonH);

    RECT frame{0, 0, s.gridRight() + pad, statusBottom + s.s(8) + buttonH + pad};
    AdjustWindowRectExForDpi(&frame, WS_POPUP | WS_CAPTION | WS_SYSMENU, FALSE, WS_EX_TOOLWINDOW | WS_EX_DLGMODALFRAME, s.dpi);
    const int w = frame.right - frame.left, h = frame.bottom - frame.top;
    MONITORINFO mi{sizeof(mi)};
    GetMonitorInfoW(MonitorFromRect(&beside, MONITOR_DEFAULTTONEAREST), &mi);
    const int x = std::clamp<int>(beside.left, mi.rcWork.left, std::max<int>(mi.rcWork.left, mi.rcWork.right - w));
    int y = beside.bottom + 8;
    if (y + h > mi.rcWork.bottom) y = std::max<int>(mi.rcWork.top, beside.top - h - 8);
    SetWindowPos(hwnd, HWND_TOP, x, y, w, h, SWP_SHOWWINDOW);
    SetFocus(hwnd);
}

void closeCalendarWindow() {
    if (g_state.window) DestroyWindow(g_state.window);
}

bool calendarWindowOpen() { return g_state.window != nullptr; }

std::wstring calendarSummary() { return g_state.window ? summaryText() : L""; }

void calendarPick(double first, double last) {
    if (!g_state.window) return;
    g_state.first = std::floor(first), g_state.last = g_state.single ? g_state.first : std::max(g_state.first, std::floor(last));
    apply();
}

}  // namespace rethink::addin
