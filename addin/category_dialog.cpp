#include "category_dialog.h"

#include <algorithm>

extern "C" IMAGE_DOS_HEADER __ImageBase;

namespace rethink::addin {
namespace {

constexpr wchar_t kClass[] = L"rethinkCategoryOrder";
enum : int { kList = 10, kUp, kDown, kOk, kCancel };

struct State {
    HWND list = nullptr;
    std::vector<std::wstring> order;
    bool done = false, accepted = false;
};
State* g_state = nullptr;

void fill(int select) {
    SendMessageW(g_state->list, LB_RESETCONTENT, 0, 0);
    for (auto& c : g_state->order) SendMessageW(g_state->list, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(c.c_str()));
    SendMessageW(g_state->list, LB_SETCURSEL, select, 0);
}

void move(int delta) {
    const int i = static_cast<int>(SendMessageW(g_state->list, LB_GETCURSEL, 0, 0));
    const int j = i + delta;
    if (i < 0 || j < 0 || j >= static_cast<int>(g_state->order.size())) return;
    std::swap(g_state->order[i], g_state->order[j]);
    fill(j);
}

LRESULT CALLBACK windowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_COMMAND:
            switch (LOWORD(wParam)) {
                case kUp: move(-1); return 0;
                case kDown: move(1); return 0;
                case kOk: g_state->accepted = true; [[fallthrough]];
                case kCancel:
                case IDCANCEL: g_state->done = true; return 0;
            }
            return 0;
        case WM_CLOSE: g_state->done = true; return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

}  // namespace

std::optional<std::vector<std::wstring>> chooseCategoryOrder(HWND owner, const std::vector<std::wstring>& categories) {
    const auto module = reinterpret_cast<HINSTANCE>(&__ImageBase);
    static bool registered = false;
    if (!registered) {
        WNDCLASSEXW wc{sizeof(wc)};
        wc.lpfnWndProc = windowProc;
        wc.hInstance = module;
        wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
        wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
        wc.lpszClassName = kClass;
        RegisterClassExW(&wc);
        registered = true;
    }
    State state;
    state.order = categories;
    g_state = &state;
    const UINT dpi = owner && GetDpiForWindow(owner) ? GetDpiForWindow(owner) : 96;
    auto s = [&](int px) { return MulDiv(px, static_cast<int>(dpi), 96); };
    HFONT font = CreateFontW(-MulDiv(9, static_cast<int>(dpi), 72), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                             CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    RECT frame{0, 0, s(330), s(250)};
    AdjustWindowRectExForDpi(&frame, WS_POPUP | WS_CAPTION | WS_SYSMENU, FALSE, WS_EX_DLGMODALFRAME, dpi);
    RECT ownerRect{};
    GetWindowRect(owner, &ownerRect);
    const int w = frame.right - frame.left, h = frame.bottom - frame.top;
    HWND hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME, kClass, L"rethink – Reihenfolge der Kategorien", WS_POPUP | WS_CAPTION | WS_SYSMENU,
                                (ownerRect.left + ownerRect.right - w) / 2, (ownerRect.top + ownerRect.bottom - h) / 2, w, h, owner, nullptr, module, nullptr);
    auto child = [&](const wchar_t* cls, const wchar_t* text, DWORD style, int x, int y, int cw, int ch, int id) {
        HWND c = CreateWindowExW(0, cls, text, WS_CHILD | WS_VISIBLE | style, s(x), s(y), s(cw), s(ch), hwnd,
                                 reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)), module, nullptr);
        SendMessageW(c, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
        return c;
    };
    child(L"STATIC", L"Die Zeilen werden in dieser Reihenfolge gruppiert:", 0, 12, 10, 300, 18, 0);
    state.list = child(L"LISTBOX", L"", WS_BORDER | WS_VSCROLL | LBS_NOTIFY | WS_TABSTOP, 12, 32, 210, 170, kList);
    child(L"BUTTON", L"Nach oben", BS_PUSHBUTTON | WS_TABSTOP, 232, 32, 86, 26, kUp);
    child(L"BUTTON", L"Nach unten", BS_PUSHBUTTON | WS_TABSTOP, 232, 64, 86, 26, kDown);
    child(L"BUTTON", L"Sortieren", BS_DEFPUSHBUTTON | WS_TABSTOP, 140, 212, 86, 28, kOk);
    child(L"BUTTON", L"Abbrechen", BS_PUSHBUTTON | WS_TABSTOP, 232, 212, 86, 28, kCancel);
    fill(0);
    if (owner) EnableWindow(owner, FALSE);
    ShowWindow(hwnd, SW_SHOW);
    SetFocus(state.list);
    // Eigene modale Schleife wie ein Dialog: Tastatur (Tab, Enter, Esc) über IsDialogMessage
    MSG msg;
    while (!state.done) {
        const BOOL got = GetMessageW(&msg, nullptr, 0, 0);
        if (got <= 0) {
            if (got == 0) PostQuitMessage(static_cast<int>(msg.wParam));  // Beenden an die äußere Schleife weitergeben
            break;
        }
        if (msg.message == WM_KEYDOWN && msg.wParam == VK_ESCAPE) state.done = true;
        else if (msg.message == WM_KEYDOWN && msg.wParam == VK_RETURN) state.done = state.accepted = true;
        else if (!IsDialogMessageW(hwnd, &msg)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
    if (owner) EnableWindow(owner, TRUE);
    DestroyWindow(hwnd);
    DeleteObject(font);
    if (owner) SetForegroundWindow(owner);
    g_state = nullptr;
    if (!state.accepted) return std::nullopt;
    return state.order;
}

}  // namespace rethink::addin
