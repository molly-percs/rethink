#include "datasheet_window.h"

#include "cellinput.h"
#include "util.h"
#include "xlsx.h"

#include <commdlg.h>
#include <windowsx.h>

#include <algorithm>
#include <format>
#include <sstream>

extern "C" IMAGE_DOS_HEADER __ImageBase;

namespace rethink::addin {
namespace {

constexpr wchar_t kClass[] = L"rethinkDatasheet";
constexpr COLORREF kGrid = RGB(225, 225, 225), kHeader = RGB(243, 243, 243), kHeaderText = RGB(68, 68, 68);
constexpr COLORREF kHeaderActive = RGB(210, 210, 210), kAccent = RGB(33, 115, 70), kSelection = RGB(232, 242, 236);
constexpr COLORREF kHint = RGB(120, 120, 120);
enum Command : UINT {
    kInsertRow = 100, kDeleteRow, kInsertColumn, kDeleteColumn, kTogglePercent, kFill, kClearFill, kExcel, kUndo, kCopy, kPaste, kSortCategories
};

}  // namespace

class Datasheet {
public:
    Datasheet(std::vector<Row> table, std::wstring separator, DatasheetCallbacks callbacks, std::wstring hint)
        : table_(std::move(table)), sep_(std::move(separator)), cb_(std::move(callbacks)), hint_(std::move(hint)) {
        if (hint_.empty()) hint_ = L"Enter übernimmt · Strg+V fügt aus Excel ein · „e“ = berechnete Summe · Rechtsklick: Zeilen, %, Farbe";
    }

    HWND create(HWND owner, const std::wstring& title, const RECT& beside, bool show) {
        static bool registered = false;
        const auto module = reinterpret_cast<HINSTANCE>(&__ImageBase);
        if (!registered) {
            WNDCLASSEXW wc{sizeof(wc)};
            wc.style = CS_DBLCLKS | CS_HREDRAW | CS_VREDRAW;
            wc.lpfnWndProc = &Datasheet::proc;
            wc.hInstance = module;
            wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
            wc.lpszClassName = kClass;
            RegisterClassExW(&wc);
            registered = true;
        }
        hwnd_ = CreateWindowExW(WS_EX_TOOLWINDOW, kClass, title.c_str(), WS_OVERLAPPEDWINDOW | WS_VSCROLL | WS_HSCROLL | WS_CLIPCHILDREN,
                                CW_USEDEFAULT, CW_USEDEFAULT, 600, 320, owner, nullptr, module, this);
        if (!hwnd_) return nullptr;
        metrics();
        place(beside);
        if (show) {
            ShowWindow(hwnd_, SW_SHOWNORMAL);
            SetForegroundWindow(hwnd_);
        }
        return hwnd_;
    }

    HWND hwnd() const { return hwnd_; }

    void setTable(std::vector<Row> table) {
        cancelEdit();
        table_ = std::move(table);
        refresh();
    }

    void activate() {
        ShowWindow(hwnd_, IsIconic(hwnd_) ? SW_RESTORE : SW_SHOWNORMAL);
        SetForegroundWindow(hwnd_);
    }

    void close() { SendMessageW(hwnd_, WM_CLOSE, 0, 0); }

    void enter(int row, int column, const std::wstring& text) {
        cur_ = anchor_ = {column, row};
        commitText(row, column, text);
    }

    void menu(int row, const std::wstring& what) {
        cur_ = anchor_ = {0, row};
        command(what == L"insertRow" ? kInsertRow : what == L"deleteRow" ? kDeleteRow : kUndo);
    }

    void paste(int row, int column, const std::wstring& text) {
        cur_ = anchor_ = {column, row};
        pasteText(text);
    }

    std::wstring text(int row, int column) const { return displayCell(get(row, column), sep_); }

private:
    struct Pos {
        int col = 0, row = 0;
    };

    // ---- Daten ----
    const Cell& get(int r, int c) const {
        static const Cell empty;
        if (r < 0 || c < 0 || r >= static_cast<int>(table_.size()) || c >= static_cast<int>(table_[r].size())) return empty;
        return table_[r][c];
    }
    void set(int r, int c, const Cell& cell) {
        if (r >= static_cast<int>(table_.size())) {
            if (cell.kind == CellKind::Empty && !cell.fill) return;
            table_.resize(r + 1);
        }
        Row& row = table_[r];
        if (c >= static_cast<int>(row.size())) {
            if (cell.kind == CellKind::Empty && !cell.fill) return;
            row.resize(c + 1);
        }
        row[c] = cell;
    }
    void normalize() {
        for (auto& r : table_)
            while (!r.empty() && r.back().kind == CellKind::Empty && !r.back().fill) r.pop_back();
        while (table_.size() > 1 && table_.back().empty()) table_.pop_back();
        if (table_.empty()) table_.push_back(Row{});
    }
    void pushUndo() {
        undo_.push_back(table_);
        undoShift_.push_back({});
        if (undo_.size() > 100) undo_.erase(undo_.begin()), undoShift_.erase(undoShift_.begin());
    }
    // Zeilen eingefügt/gelöscht: melden (Zeitplan verschiebt seine Elemente mit) und fürs Rückgängigmachen merken
    void shifted(int from, int delta) {
        undoShift_.back() = {from, delta};
        if (cb_.rowsShifted) cb_.rowsShifted(from, delta);
    }
    void changed() {
        normalize();
        refresh();
        if (cb_.changed) cb_.changed(table_);
    }
    int usedColumns() const {
        size_t n = 0;
        for (auto& r : table_) n = std::max(n, r.size());
        return static_cast<int>(n);
    }

    // ---- Geometrie ----
    void metrics() {
        const UINT dpi = GetDpiForWindow(hwnd_) ? GetDpiForWindow(hwnd_) : 96;
        cellW_ = MulDiv(88, dpi, 96), cellH_ = MulDiv(21, dpi, 96), headW_ = MulDiv(38, dpi, 96), headH_ = MulDiv(21, dpi, 96);
        hintH_ = MulDiv(20, dpi, 96);
        if (font_) DeleteObject(font_);
        font_ = CreateFontW(-MulDiv(9, dpi, 72), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                            CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    }
    void place(const RECT& beside) {
        const int cols = std::clamp(usedColumns() + 1, 5, 9);
        const int rows = std::clamp(static_cast<int>(table_.size()) + 3, 6, 14);
        RECT want{0, 0, headW_ + cols * cellW_, headH_ + rows * cellH_ + hintH_};
        AdjustWindowRectExForDpi(&want, WS_OVERLAPPEDWINDOW | WS_VSCROLL | WS_HSCROLL, FALSE, WS_EX_TOOLWINDOW, GetDpiForWindow(hwnd_));
        const int w = want.right - want.left, h = want.bottom - want.top;
        MONITORINFO mi{sizeof(mi)};
        GetMonitorInfoW(MonitorFromRect(&beside, MONITOR_DEFAULTTONEAREST), &mi);
        const RECT& work = mi.rcWork;
        int x = beside.right + 12, y = beside.top;  // bevorzugt rechts neben dem Diagramm
        if (x + w > work.right) x = std::max<int>(work.left, beside.left), y = beside.bottom + 12;  // sonst darunter
        if (y + h > work.bottom) y = std::max<int>(work.top, work.bottom - h);
        if (x + w > work.right) x = std::max<int>(work.left, work.right - w);
        SetWindowPos(hwnd_, nullptr, x, y, w, h, SWP_NOZORDER | SWP_NOACTIVATE);
    }
    RECT cellsArea() const {
        RECT r;
        GetClientRect(hwnd_, &r);
        r.left += headW_, r.top += headH_, r.bottom -= hintH_;
        return r;
    }
    int visibleRows() const {
        const RECT a = cellsArea();
        return std::max(1, static_cast<int>((a.bottom - a.top) / cellH_));
    }
    int visibleCols() const {
        const RECT a = cellsArea();
        return std::max(1, static_cast<int>((a.right - a.left) / cellW_));
    }
    int gridRows() const { return std::max(static_cast<int>(table_.size()) + 20, scrollRow_ + visibleRows() + 1); }
    int gridCols() const { return std::max(usedColumns() + 8, scrollCol_ + visibleCols() + 1); }
    RECT cellRect(int r, int c) const {
        const RECT a = cellsArea();
        const int x = a.left + (c - scrollCol_) * cellW_, y = a.top + (r - scrollRow_) * cellH_;
        return {x, y, x + cellW_, y + cellH_};
    }
    bool hit(int x, int y, Pos& p) const {
        const RECT a = cellsArea();
        if (x < a.left || y < a.top || y >= a.bottom) return false;
        p = {scrollCol_ + (x - a.left) / cellW_, scrollRow_ + (y - a.top) / cellH_};
        return true;
    }
    void updateScrollbars() {
        SCROLLINFO si{sizeof(si), SIF_ALL};
        si.nMin = 0, si.nMax = gridRows() - 1, si.nPage = visibleRows(), si.nPos = scrollRow_;
        SetScrollInfo(hwnd_, SB_VERT, &si, TRUE);
        si.nMax = gridCols() - 1, si.nPage = visibleCols(), si.nPos = scrollCol_;
        SetScrollInfo(hwnd_, SB_HORZ, &si, TRUE);
    }
    void refresh() {
        updateScrollbars();
        InvalidateRect(hwnd_, nullptr, FALSE);
    }
    void ensureVisible() {
        if (cur_.row < scrollRow_) scrollRow_ = cur_.row;
        if (cur_.row >= scrollRow_ + visibleRows()) scrollRow_ = cur_.row - visibleRows() + 1;
        if (cur_.col < scrollCol_) scrollCol_ = cur_.col;
        if (cur_.col >= scrollCol_ + visibleCols()) scrollCol_ = cur_.col - visibleCols() + 1;
        refresh();
    }
    void selection(int& r0, int& c0, int& r1, int& c1) const {
        r0 = std::min(cur_.row, anchor_.row), r1 = std::max(cur_.row, anchor_.row);
        c0 = std::min(cur_.col, anchor_.col), c1 = std::max(cur_.col, anchor_.col);
    }

    // ---- Zeichnen ----
    void paint(HDC target) {
        RECT client;
        GetClientRect(hwnd_, &client);
        HDC dc = CreateCompatibleDC(target);
        HBITMAP bmp = CreateCompatibleBitmap(target, client.right, client.bottom);
        HGDIOBJ oldBmp = SelectObject(dc, bmp), oldFont = SelectObject(dc, font_);
        SetBkMode(dc, TRANSPARENT);
        auto fill = [&](const RECT& r, COLORREF c) {
            HBRUSH b = CreateSolidBrush(c);
            FillRect(dc, &r, b);
            DeleteObject(b);
        };
        fill(client, RGB(255, 255, 255));
        int r0, c0, r1, c1;
        selection(r0, c0, r1, c1);
        const RECT area = cellsArea();
        const int rows = visibleRows() + 1, cols = visibleCols() + 1;

        // Zellen
        for (int r = scrollRow_; r < scrollRow_ + rows; ++r) {
            for (int c = scrollCol_; c < scrollCol_ + cols; ++c) {
                RECT cr = cellRect(r, c);
                if (cr.top >= area.bottom) continue;
                const Cell& cell = get(r, c);
                const bool selected = r >= r0 && r <= r1 && c >= c0 && c <= c1 && !(r == cur_.row && c == cur_.col);
                if (cell.fill) fill(cr, RGB((*cell.fill >> 16) & 0xFF, (*cell.fill >> 8) & 0xFF, *cell.fill & 0xFF));
                else if (selected) fill(cr, kSelection);
                const std::wstring s = displayCell(cell, sep_);
                if (!s.empty()) {
                    RECT tr{cr.left + 4, cr.top, cr.right - 4, cr.bottom};
                    const bool right = cell.kind == CellKind::Number || cell.kind == CellKind::Percentage || cell.kind == CellKind::Date;
                    const bool computed = cell.kind == CellKind::String && (cell.text == L"e" || cell.text == L"E");
                    SetTextColor(dc, computed ? kAccent : RGB(0, 0, 0));
                    DrawTextW(dc, s.c_str(), -1, &tr, (right || computed ? DT_RIGHT : DT_LEFT) | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
                }
            }
        }
        // Gitter
        HPEN grid = CreatePen(PS_SOLID, 1, kGrid);
        HGDIOBJ oldPen = SelectObject(dc, grid);
        for (int c = 0; c <= cols; ++c) {
            const int x = area.left + c * cellW_;
            MoveToEx(dc, x, area.top, nullptr), LineTo(dc, x, area.bottom);
        }
        for (int r = 0; r <= rows; ++r) {
            const int y = area.top + r * cellH_;
            if (y > area.bottom) break;
            MoveToEx(dc, area.left, y, nullptr), LineTo(dc, client.right, y);
        }
        // Kopfzeilen
        fill({0, 0, client.right, headH_}, kHeader);
        fill({0, 0, headW_, area.bottom}, kHeader);
        SetTextColor(dc, kHeaderText);
        for (int c = scrollCol_; c < scrollCol_ + cols; ++c) {
            RECT hr = cellRect(0, c);
            hr.top = 0, hr.bottom = headH_;
            if (c >= c0 && c <= c1) fill(hr, kHeaderActive);
            const std::wstring label = columnName(c);
            DrawTextW(dc, label.c_str(), -1, &hr, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            MoveToEx(dc, hr.right, 0, nullptr), LineTo(dc, hr.right, headH_);
        }
        for (int r = scrollRow_; r < scrollRow_ + rows; ++r) {
            RECT hr = cellRect(r, 0);
            hr.left = 0, hr.right = headW_;
            if (hr.top >= area.bottom) break;
            if (r >= r0 && r <= r1) fill(hr, kHeaderActive);
            const std::wstring label = std::to_wstring(r + 1);
            DrawTextW(dc, label.c_str(), -1, &hr, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            MoveToEx(dc, 0, hr.bottom, nullptr), LineTo(dc, headW_, hr.bottom);
        }
        SelectObject(dc, oldPen);
        DeleteObject(grid);
        // aktive Zelle
        if (cur_.row >= scrollRow_ && cur_.col >= scrollCol_) {
            RECT cr = cellRect(cur_.row, cur_.col);
            HPEN accent = CreatePen(PS_SOLID, 2, kAccent);
            HGDIOBJ op = SelectObject(dc, accent);
            HGDIOBJ ob = SelectObject(dc, GetStockObject(NULL_BRUSH));
            Rectangle(dc, cr.left, cr.top, cr.right + 1, cr.bottom + 1);
            SelectObject(dc, ob);
            SelectObject(dc, op);
            DeleteObject(accent);
        }
        // Hinweiszeile
        RECT hint{6, area.bottom, client.right - 6, client.bottom};
        fill({0, area.bottom, client.right, client.bottom}, kHeader);
        SetTextColor(dc, kHint);
        DrawTextW(dc, hint_.c_str(), -1, &hint, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

        BitBlt(target, 0, 0, client.right, client.bottom, dc, 0, 0, SRCCOPY);
        SelectObject(dc, oldFont);
        SelectObject(dc, oldBmp);
        DeleteObject(bmp);
        DeleteDC(dc);
    }

    // ---- Bearbeiten ----
    void beginEdit(const std::wstring* initial) {
        if (editing_) return;
        ensureVisible();
        const RECT cr = cellRect(cur_.row, cur_.col);
        const std::wstring text = initial ? *initial : displayCell(get(cur_.row, cur_.col), sep_);
        edit_ = CreateWindowExW(0, L"EDIT", text.c_str(), WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, cr.left + 2, cr.top + 2, cr.right - cr.left - 3,
                                cr.bottom - cr.top - 3, hwnd_, nullptr, reinterpret_cast<HINSTANCE>(&__ImageBase), nullptr);
        SendMessageW(edit_, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
        SetWindowLongPtrW(edit_, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(this));
        editProc_ = reinterpret_cast<WNDPROC>(SetWindowLongPtrW(edit_, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(&Datasheet::editProc)));
        editing_ = true;
        SetFocus(edit_);
        SendMessageW(edit_, EM_SETSEL, initial ? text.size() : 0, -1);
    }
    void cancelEdit() {
        if (!editing_) return;
        editing_ = false;
        DestroyWindow(edit_);
        edit_ = nullptr;
        if (hwnd_) SetFocus(hwnd_);
    }
    void commitEdit() {
        if (!editing_) return;
        const int length = GetWindowTextLengthW(edit_);
        std::wstring text(static_cast<size_t>(length) + 1, L'\0');
        GetWindowTextW(edit_, text.data(), length + 1);
        text.resize(static_cast<size_t>(length));
        cancelEdit();
        commitText(cur_.row, cur_.col, text);
    }
    void commitText(int r, int c, const std::wstring& text) {
        Cell cell = parseCellInput(text, sep_);
        cell.fill = get(r, c).fill;  // Zellfarbe bleibt beim Überschreiben
        const Cell& old = get(r, c);
        if (old.kind == cell.kind && old.text == cell.text && old.number == cell.number) return;
        pushUndo();
        set(r, c, cell);
        changed();
    }
    void move(int dr, int dc, bool extend) {
        cur_.row = std::max(0, cur_.row + dr);
        cur_.col = std::max(0, cur_.col + dc);
        if (!extend) anchor_ = cur_;
        ensureVisible();
    }

    // ---- Zwischenablage ----
    void copy() {
        int r0, c0, r1, c1;
        selection(r0, c0, r1, c1);
        std::wstring text;
        for (int r = r0; r <= r1; ++r) {
            for (int c = c0; c <= c1; ++c) text += (c > c0 ? L"\t" : L"") + displayCell(get(r, c), sep_);
            text += L"\r\n";
        }
        if (!OpenClipboard(hwnd_)) return;
        EmptyClipboard();
        HGLOBAL mem = GlobalAlloc(GMEM_MOVEABLE, (text.size() + 1) * sizeof(wchar_t));
        if (mem) {
            memcpy(GlobalLock(mem), text.c_str(), (text.size() + 1) * sizeof(wchar_t));
            GlobalUnlock(mem);
            SetClipboardData(CF_UNICODETEXT, mem);
        }
        CloseClipboard();
    }
    void pasteClipboard() {
        if (!OpenClipboard(hwnd_)) return;
        std::wstring text;
        if (HANDLE data = GetClipboardData(CF_UNICODETEXT)) {
            text = static_cast<const wchar_t*>(GlobalLock(data));
            GlobalUnlock(data);
        }
        CloseClipboard();
        if (!text.empty()) pasteText(text);
    }
    void pasteText(const std::wstring& text) {
        std::vector<std::vector<std::wstring>> rows;
        std::wistringstream lines(text);
        std::wstring line;
        while (std::getline(lines, line)) {
            if (!line.empty() && line.back() == L'\r') line.pop_back();
            std::vector<std::wstring> cells;
            size_t start = 0;
            for (;;) {
                const size_t tab = line.find(L'\t', start);
                cells.push_back(line.substr(start, tab == std::wstring::npos ? std::wstring::npos : tab - start));
                if (tab == std::wstring::npos) break;
                start = tab + 1;
            }
            rows.push_back(std::move(cells));
        }
        while (!rows.empty() && rows.back().size() == 1 && rows.back()[0].empty()) rows.pop_back();
        if (rows.empty()) return;
        pushUndo();
        for (size_t r = 0; r < rows.size(); ++r)
            for (size_t c = 0; c < rows[r].size(); ++c) {
                Cell cell = parseCellInput(rows[r][c], sep_);
                cell.fill = get(cur_.row + static_cast<int>(r), cur_.col + static_cast<int>(c)).fill;
                set(cur_.row + static_cast<int>(r), cur_.col + static_cast<int>(c), cell);
            }
        anchor_ = cur_;
        cur_ = {cur_.col + static_cast<int>(rows.front().size()) - 1, cur_.row + static_cast<int>(rows.size()) - 1};
        std::swap(anchor_, cur_);
        changed();
    }
    void clearSelection() {
        int r0, c0, r1, c1;
        selection(r0, c0, r1, c1);
        pushUndo();
        for (int r = r0; r <= r1; ++r)
            for (int c = c0; c <= c1; ++c) {
                Cell cell;
                cell.fill = get(r, c).fill;
                set(r, c, cell);
            }
        changed();
    }

    // ---- Kontextmenü ----
    void command(UINT id) {
        int r0, c0, r1, c1;
        selection(r0, c0, r1, c1);
        switch (id) {
            case kInsertRow:
                pushUndo();
                if (r0 <= static_cast<int>(table_.size())) {
                    table_.insert(table_.begin() + r0, Row{});
                    shifted(r0, 1);
                }
                changed();
                break;
            case kDeleteRow:
                pushUndo();
                if (const int last = std::min(r1, static_cast<int>(table_.size()) - 1); last >= r0) {
                    table_.erase(table_.begin() + r0, table_.begin() + last + 1);
                    shifted(r0, r0 - last - 1);
                }
                changed();
                break;
            case kInsertColumn:
                pushUndo();
                for (auto& row : table_)
                    if (c0 < static_cast<int>(row.size())) row.insert(row.begin() + c0, Cell{});
                changed();
                break;
            case kDeleteColumn:
                pushUndo();
                for (auto& row : table_)
                    for (int c = c1; c >= c0; --c)
                        if (c < static_cast<int>(row.size())) row.erase(row.begin() + c);
                changed();
                break;
            case kTogglePercent: {
                pushUndo();
                const bool toPercent = get(cur_.row, cur_.col).kind != CellKind::Percentage;
                for (int r = r0; r <= r1; ++r)
                    for (int c = c0; c <= c1; ++c) {
                        Cell cell = get(r, c);
                        if (cell.kind == CellKind::Number && toPercent) cell.kind = CellKind::Percentage;
                        else if (cell.kind == CellKind::Percentage && !toPercent) cell.kind = CellKind::Number;
                        else continue;
                        set(r, c, cell);
                    }
                changed();
                break;
            }
            case kFill: {
                static COLORREF custom[16] = {};
                CHOOSECOLORW cc{sizeof(cc)};
                cc.hwndOwner = hwnd_;
                cc.lpCustColors = custom;
                const auto& cell = get(cur_.row, cur_.col);
                cc.rgbResult = cell.fill ? RGB((*cell.fill >> 16) & 0xFF, (*cell.fill >> 8) & 0xFF, *cell.fill & 0xFF) : RGB(192, 0, 0);
                cc.Flags = CC_FULLOPEN | CC_RGBINIT;
                if (!ChooseColorW(&cc)) break;
                pushUndo();
                const std::uint32_t rgb = (GetRValue(cc.rgbResult) << 16) | (GetGValue(cc.rgbResult) << 8) | GetBValue(cc.rgbResult);
                for (int r = r0; r <= r1; ++r)
                    for (int c = c0; c <= c1; ++c) {
                        Cell cell2 = get(r, c);
                        cell2.fill = rgb;
                        set(r, c, cell2);
                    }
                changed();
                break;
            }
            case kClearFill:
                pushUndo();
                for (int r = r0; r <= r1; ++r)
                    for (int c = c0; c <= c1; ++c) {
                        Cell cell = get(r, c);
                        if (!cell.fill) continue;
                        cell.fill.reset();
                        set(r, c, cell);
                    }
                changed();
                break;
            case kUndo:
                if (!undo_.empty()) {
                    table_ = undo_.back();
                    undo_.pop_back();
                    // shortcut: Rückgängig nach dem Löschen bringt die Zeilen zurück, aber nicht deren Formatierung/freie Elemente
                    if (const auto [from, delta] = undoShift_.back(); delta != 0 && cb_.rowsShifted) cb_.rowsShifted(from, -delta);
                    undoShift_.pop_back();
                    changed();
                }
                break;
            case kCopy: copy(); break;
            case kPaste: pasteClipboard(); break;
            case kExcel:
                if (cb_.openInExcel) cb_.openInExcel();
                break;
            case kSortCategories:
                commitEdit();
                if (cb_.sortCategories) cb_.sortCategories();  // sortiert über das Diagramm und setzt die Tabelle neu
                break;
        }
    }
    void contextMenu(int x, int y) {
        HMENU menu = CreatePopupMenu();
        AppendMenuW(menu, MF_STRING, kCopy, L"Kopieren\tStrg+C");
        AppendMenuW(menu, MF_STRING, kPaste, L"Einfügen\tStrg+V");
        AppendMenuW(menu, MF_STRING | (undo_.empty() ? MF_GRAYED : 0), kUndo, L"Rückgängig\tStrg+Z");
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(menu, MF_STRING, kInsertRow, L"Zeile oberhalb einfügen");
        AppendMenuW(menu, MF_STRING, kDeleteRow, L"Zeile löschen");
        AppendMenuW(menu, MF_STRING, kInsertColumn, L"Spalte links einfügen");
        AppendMenuW(menu, MF_STRING, kDeleteColumn, L"Spalte löschen");
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(menu, MF_STRING, kTogglePercent, L"Prozent ein/aus");
        AppendMenuW(menu, MF_STRING, kFill, L"Zellfarbe …");
        AppendMenuW(menu, MF_STRING, kClearFill, L"Zellfarbe entfernen");
        if (cb_.sortCategories) {
            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
            AppendMenuW(menu, MF_STRING, kSortCategories, L"Nach Kategorie sortieren …");
        }
        if (cb_.openInExcel) {
            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
            AppendMenuW(menu, MF_STRING, kExcel, L"In Excel bearbeiten");
        }
        const UINT id = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, x, y, 0, hwnd_, nullptr);
        DestroyMenu(menu);
        if (id) command(id);
    }

    // ---- Nachrichten ----
    LRESULT handle(UINT msg, WPARAM wParam, LPARAM lParam) {
        switch (msg) {
            case WM_PAINT: {
                PAINTSTRUCT ps;
                HDC dc = BeginPaint(hwnd_, &ps);
                paint(dc);
                EndPaint(hwnd_, &ps);
                return 0;
            }
            case WM_ERASEBKGND: return 1;
            case WM_SIZE: refresh(); return 0;
            case WM_DPICHANGED: {
                metrics();
                const RECT* r = reinterpret_cast<const RECT*>(lParam);
                SetWindowPos(hwnd_, nullptr, r->left, r->top, r->right - r->left, r->bottom - r->top, SWP_NOZORDER | SWP_NOACTIVATE);
                return 0;
            }
            case WM_LBUTTONDOWN: {
                commitEdit();
                SetFocus(hwnd_);
                Pos p;
                if (hit(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), p)) {
                    cur_ = p;
                    if (!(wParam & MK_SHIFT)) anchor_ = p;
                    dragging_ = true;
                    SetCapture(hwnd_);
                    refresh();
                }
                return 0;
            }
            case WM_MOUSEMOVE: {
                Pos p;
                if (dragging_ && hit(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), p) && (p.row != cur_.row || p.col != cur_.col)) {
                    cur_ = p;
                    refresh();
                }
                return 0;
            }
            case WM_LBUTTONUP:
                dragging_ = false;
                ReleaseCapture();
                return 0;
            case WM_LBUTTONDBLCLK: {
                Pos p;
                if (hit(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), p)) {
                    cur_ = anchor_ = p;
                    beginEdit(nullptr);
                }
                return 0;
            }
            case WM_RBUTTONDOWN: {
                commitEdit();
                Pos p;
                if (hit(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), p)) {
                    int r0, c0, r1, c1;
                    selection(r0, c0, r1, c1);
                    if (p.row < r0 || p.row > r1 || p.col < c0 || p.col > c1) cur_ = anchor_ = p;
                    refresh();
                }
                POINT screen{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
                ClientToScreen(hwnd_, &screen);
                contextMenu(screen.x, screen.y);
                return 0;
            }
            case WM_MOUSEWHEEL:
                scrollRow_ = std::max(0, scrollRow_ - GET_WHEEL_DELTA_WPARAM(wParam) / WHEEL_DELTA * 3);
                refresh();
                return 0;
            case WM_VSCROLL:
            case WM_HSCROLL: {
                const bool vertical = msg == WM_VSCROLL;
                int& pos = vertical ? scrollRow_ : scrollCol_;
                SCROLLINFO si{sizeof(si), SIF_ALL};
                GetScrollInfo(hwnd_, vertical ? SB_VERT : SB_HORZ, &si);
                switch (LOWORD(wParam)) {
                    case SB_LINEUP: --pos; break;
                    case SB_LINEDOWN: ++pos; break;
                    case SB_PAGEUP: pos -= static_cast<int>(si.nPage); break;
                    case SB_PAGEDOWN: pos += static_cast<int>(si.nPage); break;
                    case SB_THUMBTRACK:
                    case SB_THUMBPOSITION: pos = si.nTrackPos; break;
                }
                pos = std::max(0, pos);
                refresh();
                return 0;
            }
            case WM_KEYDOWN: {
                const bool shift = GetKeyState(VK_SHIFT) < 0, ctrl = GetKeyState(VK_CONTROL) < 0;
                if (ctrl) {
                    switch (wParam) {
                        case 'C': copy(); return 0;
                        case 'V': pasteClipboard(); return 0;
                        case 'X': copy(); clearSelection(); return 0;
                        case 'Z': command(kUndo); return 0;
                    }
                }
                switch (wParam) {
                    case VK_UP: move(-1, 0, shift); return 0;
                    case VK_DOWN: move(1, 0, shift); return 0;
                    case VK_LEFT: move(0, -1, shift); return 0;
                    case VK_RIGHT: move(0, 1, shift); return 0;
                    case VK_RETURN: move(shift ? -1 : 1, 0, false); return 0;
                    case VK_TAB: move(0, shift ? -1 : 1, false); return 0;
                    case VK_HOME: cur_.col = 0, anchor_ = cur_, ensureVisible(); return 0;
                    case VK_DELETE: clearSelection(); return 0;
                    case VK_F2: beginEdit(nullptr); return 0;
                    case VK_ESCAPE: return 0;
                }
                break;
            }
            case WM_CHAR:
                if (wParam >= 32 && GetKeyState(VK_CONTROL) >= 0) {
                    const std::wstring initial(1, static_cast<wchar_t>(wParam));
                    anchor_ = cur_;
                    beginEdit(&initial);
                    return 0;
                }
                return 0;
            case WM_GETDLGCODE: return DLGC_WANTALLKEYS;
            case WM_CLOSE:
                commitEdit();
                DestroyWindow(hwnd_);
                return 0;
            case WM_NCDESTROY: {
                auto callback = cb_.closed;
                if (font_) DeleteObject(font_);
                hwnd_ = nullptr;
                delete this;
                if (callback) callback();
                return 0;
            }
        }
        return DefWindowProcW(hwnd_, msg, wParam, lParam);
    }

    static LRESULT CALLBACK proc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
        if (msg == WM_NCCREATE) {
            auto* self = static_cast<Datasheet*>(reinterpret_cast<CREATESTRUCTW*>(lParam)->lpCreateParams);
            self->hwnd_ = hwnd;
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
        auto* self = reinterpret_cast<Datasheet*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (!self) return DefWindowProcW(hwnd, msg, wParam, lParam);
        try {
            return self->handle(msg, wParam, lParam);
        } catch (...) {
            return 0;  // nie eine Ausnahme in PowerPoints Nachrichtenschleife
        }
    }

    static LRESULT CALLBACK editProc(HWND edit, UINT msg, WPARAM wParam, LPARAM lParam) {
        auto* self = reinterpret_cast<Datasheet*>(GetWindowLongPtrW(edit, GWLP_USERDATA));
        if (msg == WM_GETDLGCODE) return DLGC_WANTALLKEYS;
        if (msg == WM_KEYDOWN && self) {
            const bool shift = GetKeyState(VK_SHIFT) < 0;
            switch (wParam) {
                case VK_RETURN: self->commitEdit(); self->move(shift ? -1 : 1, 0, false); return 0;
                case VK_TAB: self->commitEdit(); self->move(0, shift ? -1 : 1, false); return 0;
                case VK_ESCAPE: self->cancelEdit(); return 0;
                case VK_UP: self->commitEdit(); self->move(-1, 0, false); return 0;
                case VK_DOWN: self->commitEdit(); self->move(1, 0, false); return 0;
            }
        }
        if (msg == WM_CHAR && (wParam == VK_RETURN || wParam == VK_TAB || wParam == VK_ESCAPE)) return 0;  // kein Piepton
        if (msg == WM_KILLFOCUS && self && self->editing_) {
            const LRESULT r = CallWindowProcW(self->editProc_, edit, msg, wParam, lParam);
            self->commitEdit();
            return r;
        }
        return self ? CallWindowProcW(self->editProc_, edit, msg, wParam, lParam) : DefWindowProcW(edit, msg, wParam, lParam);
    }

    HWND hwnd_ = nullptr, edit_ = nullptr;
    WNDPROC editProc_ = nullptr;
    HFONT font_ = nullptr;
    std::vector<Row> table_;
    std::vector<std::vector<Row>> undo_;
    std::vector<std::pair<int, int>> undoShift_;  // je Rückgängig-Stand: eingefügte/gelöschte Zeilen (from, delta)
    std::wstring sep_;
    DatasheetCallbacks cb_;
    std::wstring hint_;
    Pos cur_{1, 1}, anchor_{1, 1};
    int scrollRow_ = 0, scrollCol_ = 0;
    int cellW_ = 88, cellH_ = 21, headW_ = 38, headH_ = 21, hintH_ = 20;
    bool editing_ = false, dragging_ = false;
};

Datasheet* openDatasheetWindow(HWND owner, const std::wstring& title, std::vector<Row> table, const std::wstring& decimalSeparator,
                               DatasheetCallbacks callbacks, const RECT& beside, bool show, const std::wstring& hint) {
    auto* sheet = new Datasheet(std::move(table), decimalSeparator, std::move(callbacks), hint);
    if (!sheet->create(owner, title, beside, show)) {
        delete sheet;
        throw Error("Das Datenblatt-Fenster konnte nicht geöffnet werden.");
    }
    return sheet;
}

void datasheetSetTable(Datasheet* sheet, std::vector<Row> table) { sheet->setTable(std::move(table)); }
void datasheetActivate(Datasheet* sheet) { sheet->activate(); }
void datasheetClose(Datasheet* sheet) { sheet->close(); }
HWND datasheetWindow(Datasheet* sheet) { return sheet->hwnd(); }
void datasheetEnter(Datasheet* sheet, int row, int column, const std::wstring& text) { sheet->enter(row, column, text); }
void datasheetPaste(Datasheet* sheet, int row, int column, const std::wstring& text) { sheet->paste(row, column, text); }
void datasheetMenu(Datasheet* sheet, int row, const std::wstring& what) { sheet->menu(row, what); }
std::wstring datasheetText(Datasheet* sheet, int row, int column) { return sheet->text(row, column); }

}  // namespace rethink::addin
