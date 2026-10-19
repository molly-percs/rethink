#include <winrt/base.h>

#include "gantt_menu.h"

#include "dispatch.h"
#include "engine/chart.h"
#include "engine/dates.h"

#include <commdlg.h>

#include <algorithm>
#include <cmath>
#include <format>
#include <vector>

namespace rethink::addin {
namespace {

using namespace engine;

struct Entry {
    std::wstring label, command;
    bool checked = false, radio = false, enabled = true, separator = false;
    std::vector<Entry> sub;
    long swatch = -1;  // Farbfeld vor dem Eintrag (RGB wie PowerPoint), -1 = keins
};

Entry item(std::wstring label, std::wstring command, bool checked = false, bool radio = false, bool enabled = true) {
    return {std::move(label), std::move(command), checked, radio, enabled, false, {}};
}
Entry toggle(std::wstring label, const std::wstring& key, bool on) { return item(std::move(label), key + (on ? L"=0" : L"=1"), on); }
Entry submenu(std::wstring label, std::vector<Entry> entries, bool enabled = true) {
    Entry e;
    e.label = std::move(label), e.sub = std::move(entries), e.enabled = enabled;
    return e;
}
Entry separator() {
    Entry e;
    e.separator = true;
    return e;
}

const wchar_t* unitLabel(TimeUnit u) {
    switch (u) {
        case TimeUnit::Year: return L"Jahr";
        case TimeUnit::Quarter: return L"Quartal";
        case TimeUnit::Month: return L"Monat";
        case TimeUnit::Week: return L"Woche";
        case TimeUnit::Day: return L"Tag";
    }
    return L"";
}

const wchar_t* unitKey(TimeUnit u) {
    switch (u) {
        case TimeUnit::Year: return L"year";
        case TimeUnit::Quarter: return L"quarter";
        case TimeUnit::Month: return L"month";
        case TimeUnit::Week: return L"week";
        case TimeUnit::Day: return L"day";
    }
    return L"";
}

const std::wstring& storedFormat(const GanttOptions& g, TimeUnit u) {
    switch (u) {
        case TimeUnit::Year: return g.yearFormat;
        case TimeUnit::Quarter: return g.quarterFormat;
        case TimeUnit::Month: return g.monthFormat;
        case TimeUnit::Week: return g.weekFormat;
        case TimeUnit::Day: return g.dayFormat;
    }
    return g.dayFormat;
}

// Formate je Skala (Codes wie im Original dokumentiert); das Menü zeigt sie mit dem heutigen Datum als Beispiel
std::vector<const wchar_t*> scaleFormats(TimeUnit u) {
    switch (u) {
        case TimeUnit::Year: return {L"yyyy", L"yy", L"'yy"};
        case TimeUnit::Quarter: return {L"\"Q\"Q", L"\"Q\"Q yyyy", L"QQQ", L"Q\". Quartal\""};
        case TimeUnit::Month: return {L"MMMM", L"MMM", L"MMMMM", L"MM", L"MMM yy", L"MMMM yyyy"};
        case TimeUnit::Week: return {L"www", L"ww", L"w", L"dd.MM.", L"dd.MM.yy"};
        case TimeUnit::Day: return {L"d", L"dd", L"d.", L"ddd", L"ddddd", L"ddd d"};
    }
    return {};
}

std::wstring shapeLabel(const std::wstring& s) {
    if (s == L"bar") return L"Balken";
    if (s == L"rounded") return L"Abgerundet";
    if (s == L"pentagon") return L"Fünfeck (Prozesspfeil)";
    if (s == L"chevron") return L"Chevron";
    if (s == L"diamond") return L"Raute";
    if (s == L"triangle") return L"Dreieck";
    if (s == L"circle") return L"Kreis";
    if (s == L"square") return L"Quadrat";
    if (s == L"star") return L"Stern";
    return s;
}

const std::vector<std::wstring> kBarShapes = {L"bar", L"rounded", L"pentagon", L"chevron"};
const std::vector<std::wstring> kMarkers = {L"diamond", L"triangle", L"circle", L"square", L"star"};

std::vector<Entry> barShapeEntries(const GanttState& s) {
    std::vector<Entry> shapes;
    if (!s.targetIsItem) shapes.push_back(item(L"Wie Vorgabe (" + shapeLabel(s.barShape) + L")", L"rowShape=", s.rowShape.empty(), true));
    for (auto& sh : kBarShapes)
        shapes.push_back(item(shapeLabel(sh), L"rowShape=" + sh, (s.rowShape.empty() ? (s.targetIsItem ? L"bar" : L"") : s.rowShape) == sh, true));
    return shapes;
}

std::vector<Entry> markerEntries(const GanttState& s) {
    std::vector<Entry> shapes;
    if (!s.targetIsItem) {
        shapes.push_back(item(L"Wie Vorgabe (" + shapeLabel(s.marker) + L")", L"rowMarker=", s.rowMarker.empty(), true));
        for (auto& sh : kMarkers) shapes.push_back(item(shapeLabel(sh), L"rowMarker=" + sh, s.rowMarker == sh, true));
    } else {
        for (auto& sh : kMarkers) shapes.push_back(item(shapeLabel(sh), L"milestoneShape=" + sh, s.marker == sh, true));
    }
    return shapes;
}

std::vector<Entry> scaleEntries(const GanttState& s) {
    std::vector<Entry> scales;
    for (TimeUnit u : {TimeUnit::Year, TimeUnit::Quarter, TimeUnit::Month, TimeUnit::Week, TimeUnit::Day}) {
        const bool shown = std::find(s.scales.begin(), s.scales.end(), u) != s.scales.end();
        scales.push_back(toggle(unitLabel(u), std::wstring(L"scale:") + unitKey(u), shown));
    }
    scales.push_back(separator());
    scales.push_back(item(L"Automatisch (nach Platz)", L"scales=auto", s.options.scales.empty()));
    return scales;
}

std::vector<Entry> formatEntries(const GanttState& s) {
    std::vector<Entry> formats;
    for (TimeUnit u : s.scales) {
        const std::wstring& current = storedFormat(s.options, u);
        std::vector<Entry> list{item(L"Automatisch (nach Platz)", std::format(L"format:{}=", unitKey(u)), current.empty(), true)};
        for (const wchar_t* code : scaleFormats(u)) {
            std::wstring example = formatDate(s.today, code, s.language);
            for (auto& c : example)
                if (c == L'\n') c = L' ';
            list.push_back(item(example, std::format(L"format:{}={}", unitKey(u), code), current == code, true));
        }
        formats.push_back(submenu(unitLabel(u), list));
    }
    return formats;
}

const wchar_t* const kDateFormats[] = {L"dd.MM.", L"dd.MM.yy", L"dd.MM.yyyy", L"d. MMM", L"d. MMM yy", L"ddd dd.MM.", L"MMM yy"};

std::wstring currentDateFormat(const GanttState& s) { return s.options.dateFormat.empty() ? L"dd.MM." : s.options.dateFormat; }

// Datumsformate mit dem heutigen Datum als Beispiel
std::vector<Entry> dateFormatEntries(const GanttState& s) {
    std::vector<Entry> list;
    for (const wchar_t* code : kDateFormats)
        list.push_back(item(formatDate(s.today, code, s.language), std::wstring(L"dateFormat=") + code, currentDateFormat(s) == code, true));
    return list;
}

std::vector<Entry> barLabelEntries(const GanttOptions& g) {
    return {item(L"Keins", L"barLabels=none", g.barLabels == BarLabel::None, true), item(L"Zeitraum", L"barLabels=range", g.barLabels == BarLabel::Range, true),
            item(L"Anfang", L"barLabels=start", g.barLabels == BarLabel::Start, true), item(L"Ende", L"barLabels=end", g.barLabels == BarLabel::End, true)};
}

std::vector<Entry> buildMenu(const GanttState& s, const GanttHit& hit) {
    const GanttOptions& g = s.options;
    const bool onRow = hit.row >= gantt::kFirstActivityRow;
    std::vector<Entry> m;

    // Element unter der Maus
    if (s.targetKind == "bar") {
        m.push_back(submenu(L"Form", barShapeEntries(s)));
        if (!s.targetIsItem && onRow) {
            std::vector<Entry> links;
            for (size_t i = 0; i < s.activities.size(); ++i) {
                const std::wstring& name = s.activities[i];
                if (name.empty() || static_cast<int>(i) + gantt::kFirstActivityRow == hit.row) continue;
                const bool linked = std::find(s.rowLinks.begin(), s.rowLinks.end(), name) != s.rowLinks.end();
                links.push_back(item(name, L"link=" + name, linked));
            }
            m.push_back(submenu(L"Verbinden mit (Vorgänger)", links, !links.empty()));
        }
    } else if (s.targetKind == "milestone") {
        m.push_back(submenu(L"Form", markerEntries(s)));
    }
    if (!s.targetKind.empty()) {
        m.push_back(item(L"Kalender öffnen …", L"calendar"));
        m.push_back(item(L"Element löschen", L"deleteItem"));
        m.push_back(separator());
    }

    // Zeile und neue Elemente an der Mausposition
    if (onRow) {
        m.push_back(submenu(L"Zeile", {item(L"Zeile einfügen", L"insertRow"), item(L"Zeile duplizieren", L"duplicateRow"),
                                        item(L"Zeile löschen", L"deleteRow"), separator(), item(L"Nach oben", L"rowUp"), item(L"Nach unten", L"rowDown"),
                                        separator(), toggle(L"Zeile schattieren", L"rowShade", s.rowShaded)}));
    } else {
        m.push_back(item(L"Zeile anfügen", L"insertRow"));
    }
    if (hit.hasDate && !hit.inScale) {
        m.push_back(submenu(L"Neu an dieser Stelle", {item(L"Balken", L"new=bar", false, false, onRow), item(L"Meilenstein", L"new=milestone"),
                                                       item(L"Klammer", L"new=bracket"), item(L"Schattierung", L"new=shade"),
                                                       item(L"Fünfeck (Prozesspfeil)", L"new=pentagon", false, false, onRow),
                                                       item(L"Chevron", L"new=chevron", false, false, onRow)}));
    }
    m.push_back(separator());

    // Skalen und ihr Format
    m.push_back(submenu(L"Skalen", scaleEntries(s)));
    const auto formats = formatEntries(s);
    m.push_back(submenu(L"Format der Skalen", formats, !formats.empty()));

    // Linien und Wochenenden
    m.push_back(submenu(L"Trennlinien", {toggle(L"Senkrecht (an den Skaleneinheiten)", L"separators", s.separators),
                                         toggle(L"Waagerecht (zwischen den Zeilen)", L"rowSeparators", s.rowSeparators), separator(),
                                         toggle(L"Meilensteinlinien", L"milestoneLines", s.milestoneLines), toggle(L"Heute-Linie", L"today", g.todayLine)}));
    m.push_back(submenu(L"Wochenenden", {item(L"Normal", L"weekends=plain", g.weekends == WeekendMode::Plain, true),
                                         item(L"Leicht grau", L"weekends=shaded", g.weekends == WeekendMode::Shaded, true),
                                         item(L"Ausgeblendet", L"weekends=hidden", g.weekends == WeekendMode::Hidden, true)}));

    // Beschriftung
    m.push_back(submenu(L"Beschriftung",
                        {submenu(L"Datum an Balken", barLabelEntries(g)), toggle(L"Datum an Meilensteinen", L"milestoneLabels", g.milestoneLabels),
                         submenu(L"Datumsformat", dateFormatEntries(s)), separator(),
                         toggle(L"Spaltenüberschriften", L"headlines", g.headlines)}));
    m.push_back(submenu(L"Spalten", {toggle(L"Kategorie", L"column:category", s.categoryShown), toggle(L"Verantwortlich", L"column:responsible", s.responsibleShown),
                                     toggle(L"Bemerkung", L"column:remark", s.remarkShown)}));

    // Balken und Meilensteine
    std::vector<Entry> heights;
    const std::pair<const wchar_t*, double> kHeights[] = {{L"Schmal", 0.35}, {L"Mittel", 0.55}, {L"Breit", 0.75}, {L"Sehr breit", 0.9}};
    double nearest = 9;
    for (auto& [label, h] : kHeights) nearest = std::fabs(h - s.barHeight) < std::fabs(nearest - s.barHeight) ? h : nearest;
    for (auto& [label, h] : kHeights) heights.push_back(item(label, std::format(L"barHeight={:.2f}", h), h == nearest, true));
    std::vector<Entry> barShapes, markers;
    for (auto& sh : kBarShapes) barShapes.push_back(item(shapeLabel(sh), L"barShape=" + sh, s.barShape == sh, true));
    for (auto& sh : kMarkers) markers.push_back(item(shapeLabel(sh), L"milestoneShape=" + sh, s.marker == sh, true));
    heights.push_back(separator());
    heights.push_back(submenu(L"Form aller Balken", barShapes));
    heights.push_back(submenu(L"Form aller Meilensteine", markers));
    m.push_back(submenu(L"Balkenbreite und Formen", heights));

    // Kategorien, Verbindungen, Design
    m.push_back(submenu(L"Kategorien", {item(L"Klammer", L"categoryStyle=bracket", s.categoryStyle == CategoryStyle::Bracket, true),
                                        item(L"Farbiger Block", L"categoryStyle=block", s.categoryStyle == CategoryStyle::Block, true),
                                        item(L"Nur Text", L"categoryStyle=text", s.categoryStyle == CategoryStyle::Text, true), separator(),
                                        toggle(L"Balken in Kategoriefarbe", L"categoryColors", s.categoryColors), separator(),
                                        item(L"Reihenfolge wählen und sortieren …", L"sortCategories", false, false, !s.categories.empty())}));
    m.push_back(toggle(L"Verbindungen (Vorgänger) zeigen", L"links", g.links));
    m.push_back(submenu(L"Design", {item(L"Standard", L"design=standard", g.design == GanttDesign::Standard, true),
                                    item(L"Modern", L"design=modern", g.design == GanttDesign::Modern, true),
                                    item(L"Kontrast", L"design=contrast", g.design == GanttDesign::Contrast, true),
                                    item(L"Minimal", L"design=minimal", g.design == GanttDesign::Minimal, true)}));
    m.push_back(separator());
    m.push_back(item(L"Zeitraum (Kalender) …", L"calendar=range"));
    m.push_back(item(L"Skala an Daten anpassen", L"fit"));
    m.push_back(item(L"Datenblatt öffnen", L"datasheet"));
    return m;
}

int px(int v) { return MulDiv(v, static_cast<int>(GetDpiForSystem()), 96); }

// Farbfeld als Menübild (Rahmen grau); rgb < 0: weiß mit rotem Schrägstrich („keine Farbe“)
HBITMAP swatchBitmap(long rgb, int size) {
    HDC screen = GetDC(nullptr), dc = CreateCompatibleDC(screen);
    HBITMAP bitmap = CreateCompatibleBitmap(screen, size, size);
    HGDIOBJ old = SelectObject(dc, bitmap);
    RECT r{0, 0, size, size};
    HBRUSH brush = CreateSolidBrush(rgb >= 0 ? static_cast<COLORREF>(rgb) : RGB(255, 255, 255));
    FillRect(dc, &r, brush);
    DeleteObject(brush);
    if (rgb < 0) {
        HPEN pen = CreatePen(PS_SOLID, std::max(1, size / 8), RGB(192, 0, 0));
        HGDIOBJ oldPen = SelectObject(dc, pen);
        MoveToEx(dc, 1, size - 2, nullptr);
        LineTo(dc, size - 1, 0);
        SelectObject(dc, oldPen);
        DeleteObject(pen);
    }
    FrameRect(dc, &r, static_cast<HBRUSH>(GetStockObject(GRAY_BRUSH)));
    SelectObject(dc, old);
    DeleteDC(dc);
    ReleaseDC(nullptr, screen);
    return bitmap;
}

HMENU toMenu(const std::vector<Entry>& entries, std::vector<std::wstring>& commands, std::vector<HBITMAP>& bitmaps) {
    HMENU menu = CreatePopupMenu();
    UINT position = 0;
    for (auto& e : entries) {
        MENUITEMINFOW info{sizeof(info)};
        if (e.separator) {
            info.fMask = MIIM_FTYPE;
            info.fType = MFT_SEPARATOR;
        } else {
            info.fMask = MIIM_STRING | MIIM_STATE | MIIM_FTYPE | MIIM_ID;
            info.fType = MFT_STRING | (e.radio ? MFT_RADIOCHECK : 0);
            info.fState = (e.checked ? MFS_CHECKED : 0) | (e.enabled ? MFS_ENABLED : MFS_DISABLED);
            info.dwTypeData = const_cast<LPWSTR>(e.label.c_str());
            if (e.swatch >= 0) {
                info.fMask |= MIIM_BITMAP;
                info.hbmpItem = bitmaps.emplace_back(swatchBitmap(e.swatch, px(14)));
            }
            if (!e.sub.empty()) {
                info.fMask |= MIIM_SUBMENU;
                info.hSubMenu = toMenu(e.sub, commands, bitmaps);
            } else {
                commands.push_back(e.command);
                info.wID = static_cast<UINT>(commands.size());
            }
        }
        InsertMenuItemW(menu, position++, TRUE, &info);
    }
    return menu;
}

// Menü zeigen und den gewählten Befehl liefern (leer = abgebrochen)
std::wstring trackMenu(const std::vector<Entry>& entries, HWND owner, POINT at, UINT flags) {
    std::vector<std::wstring> commands;
    std::vector<HBITMAP> bitmaps;
    HMENU menu = toMenu(entries, commands, bitmaps);
    const UINT id = static_cast<UINT>(TrackPopupMenuEx(menu, flags | TPM_RETURNCMD | TPM_NONOTIFY, at.x, at.y, owner, nullptr));
    DestroyMenu(menu);
    for (HBITMAP b : bitmaps) DeleteObject(b);
    return id == 0 || id > commands.size() ? std::wstring() : commands[id - 1];
}

void toText(const std::vector<Entry>& entries, const std::wstring& path, std::wstring& out) {
    for (auto& e : entries) {
        if (e.separator) continue;
        const std::wstring name = path.empty() ? e.label : path + L" > " + e.label;
        if (!e.sub.empty()) {
            toText(e.sub, name, out);
            continue;
        }
        out += name + (e.checked ? L" [x]" : L"") + (e.enabled ? L"" : L" (aus)") + L"=" + e.command + L"\n";
    }
}

// ---- Minileiste: Aussehen der Stelle unter der Maus, über dem Kontextmenü (wie im Original) ----

struct Tool {
    std::wstring label;
    long swatch = -2;  // Farbfeld auf der Schaltfläche: -2 = keins, -1 = „keine Farbe“
    std::vector<Entry> menu;
};

const wchar_t* themeLabel(const std::wstring& name) {
    static const std::pair<const wchar_t*, const wchar_t*> kLabels[] = {
        {L"tx1", L"Text 1"},      {L"bg1", L"Hintergrund 1"}, {L"tx2", L"Text 2"},      {L"bg2", L"Hintergrund 2"}, {L"accent1", L"Akzent 1"},
        {L"accent2", L"Akzent 2"}, {L"accent3", L"Akzent 3"},  {L"accent4", L"Akzent 4"}, {L"accent5", L"Akzent 5"},  {L"accent6", L"Akzent 6"}};
    for (auto& [key, label] : kLabels)
        if (name == key) return label;
    return L"";
}

// Farben aus dem Design des Folienmasters; „Weitere Farben …“ öffnet die Farbauswahl (Befehl endet auf "=?")
// automatic: Vorgabefarbe der Form (Designfarbe des Masters), als Farbfeld neben „Automatisch“
std::vector<Entry> colorEntries(const GanttState& s, const std::wstring& key, const std::wstring& current, const std::wstring& automatic) {
    Entry standard = item(L"Automatisch (Master)", key + L"=", current.empty(), true);
    for (auto& [name, rgb] : s.theme)
        if (name == automatic) standard.swatch = rgb;
    std::vector<Entry> m{standard, item(L"Keine", key + L"=none", current == L"none", true), separator()};
    for (auto& [name, rgb] : s.theme) {
        Entry e = item(themeLabel(name), key + L"=" + name, current == name, true);
        e.swatch = rgb;
        m.push_back(std::move(e));
    }
    m.push_back(separator());
    m.push_back(item(L"Weitere Farben …", key + L"=?", current.starts_with(L"#")));
    return m;
}

std::wstring points(double w) {
    std::wstring t = std::format(L"{:g} pt", w);
    for (auto& c : t)
        if (c == L'.') c = L',';
    return t;
}

std::vector<Entry> weightEntries(const GanttState& s) {
    std::vector<Entry> m;
    for (double w : {0.25, 0.5, 0.75, 1.0, 1.5, 2.25, 3.0, 4.5, 6.0})
        m.push_back(item(points(w), std::format(L"lineWeight={:g}", w), s.targetLine >= 0 && std::fabs(s.targetWeight - w) < 0.01, true));
    m.push_back(separator());
    m.push_back(item(L"Durchgezogen", L"dashed=0", !s.targetStyle.dashed, true));
    m.push_back(item(L"Gestrichelt", L"dashed=1", s.targetStyle.dashed, true));
    return m;
}

std::vector<Tool> buildToolbar(const GanttState& s, const GanttHit& hit) {
    const ShapeStyle& st = s.targetStyle;
    const GanttOptions& g = s.options;
    auto fill = [&] { return Tool{L"Füllung", s.targetFill, colorEntries(s, L"fill", st.fill, s.defaultFill)}; };
    auto line = [&] { return Tool{L"Linie", s.targetLine, colorEntries(s, L"line", st.line, s.targetKind == "bracket" ? s.defaultFill : L"tx1")}; };
    auto weight = [&] { return Tool{s.targetLine >= 0 ? points(s.targetWeight) : L"Linienstärke", -2, weightEntries(s)}; };
    auto dateFormat = [&] { return Tool{formatDate(s.today, currentDateFormat(s), s.language), -2, dateFormatEntries(s)}; };
    std::vector<Tool> t;
    if (s.targetKind == "bar") {
        std::vector<Entry> form{item(L"Balken", L"barStyle=solid", !st.dashed, true), item(L"Gestrichelter Balken", L"barStyle=dashed", st.dashed, true),
                                separator()};
        for (auto& e : barShapeEntries(s)) form.push_back(e);
        t.push_back({st.dashed ? L"Gestrichelter Balken" : L"Balken", -2, form});
        t.push_back(fill()), t.push_back(line()), t.push_back(weight());
    } else if (s.targetKind == "milestone") {
        t.push_back({shapeLabel(s.rowMarker.empty() ? s.marker : s.rowMarker), -2, markerEntries(s)});
        t.push_back(fill()), t.push_back(line());
        t.push_back({L"Datum", -2, {toggle(L"Datum zeigen", L"milestoneLabels", g.milestoneLabels)}});
    } else if (s.targetKind == "bracket") {
        t.push_back(line()), t.push_back(weight());
    } else if (s.targetKind == "shade") {
        t.push_back(fill());
    } else if (!hit.label.empty()) {
        if (hit.label.starts_with("milestone"))
            t.push_back({L"Datum", -2, {toggle(L"Datum zeigen", L"milestoneLabels", g.milestoneLabels)}});
        else
            t.push_back({barLabelEntries(g)[static_cast<int>(g.barLabels)].label, -2, barLabelEntries(g)});
        t.push_back(dateFormat());
    } else if (hit.inScale) {
        t.push_back({L"Skalen", -2, scaleEntries(s)});
        t.push_back({L"Format", -2, formatEntries(s)});
    }
    return t;
}

struct MiniBar {
    HWND hwnd = nullptr;
    ComPtr<IDispatch> app;
    GanttHit hit;
    std::vector<Tool> tools;
    std::vector<RECT> rects;
    int hot = -1;
    bool busy = false;  // eigenes Menü oder Farbauswahl offen: nicht schließen
    HFONT font = nullptr;
};
MiniBar g_bar;
constexpr wchar_t kBarClass[] = L"rethinkGanttMiniBar";
constexpr UINT_PTR kBarTimer = 1;

SIZE layoutBar() {
    HDC dc = GetDC(nullptr);
    HGDIOBJ old = SelectObject(dc, g_bar.font);
    const int top = px(3), h = px(24);
    int x = px(3);
    g_bar.rects.clear();
    for (auto& t : g_bar.tools) {
        SIZE text{};
        GetTextExtentPoint32W(dc, t.label.c_str(), static_cast<int>(t.label.size()), &text);
        const int w = px(8) + (t.swatch != -2 ? px(20) : 0) + text.cx + px(16);
        g_bar.rects.push_back({x, top, x + w, top + h});
        x += w + px(2);
    }
    SelectObject(dc, old);
    ReleaseDC(nullptr, dc);
    return {x + px(1), top + h + px(3)};
}

void paintBar(HWND hwnd) {
    PAINTSTRUCT ps;
    HDC dc = BeginPaint(hwnd, &ps);
    RECT client;
    GetClientRect(hwnd, &client);
    FillRect(dc, &client, GetSysColorBrush(COLOR_MENU));
    HGDIOBJ old = SelectObject(dc, g_bar.font);
    SetBkMode(dc, TRANSPARENT);
    SetTextColor(dc, GetSysColor(COLOR_MENUTEXT));
    for (size_t i = 0; i < g_bar.tools.size(); ++i) {
        const Tool& t = g_bar.tools[i];
        RECT r = g_bar.rects[i];
        if (static_cast<int>(i) == g_bar.hot) {
            HBRUSH hot = CreateSolidBrush(RGB(229, 241, 251)), frame = CreateSolidBrush(RGB(204, 228, 247));
            FillRect(dc, &r, hot);
            FrameRect(dc, &r, frame);
            DeleteObject(hot);
            DeleteObject(frame);
        }
        int x = r.left + px(6);
        if (t.swatch != -2) {
            const int size = px(14), y = (r.top + r.bottom - size) / 2;
            HBITMAP b = swatchBitmap(t.swatch, size);
            HDC mem = CreateCompatibleDC(dc);
            HGDIOBJ oldBitmap = SelectObject(mem, b);
            BitBlt(dc, x, y, size, size, mem, 0, 0, SRCCOPY);
            SelectObject(mem, oldBitmap);
            DeleteDC(mem);
            DeleteObject(b);
            x += px(20);
        }
        RECT text{x, r.top, r.right - px(14), r.bottom};
        DrawTextW(dc, t.label.c_str(), -1, &text, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_NOPREFIX);
        RECT arrow{r.right - px(14), r.top, r.right - px(2), r.bottom};
        DrawTextW(dc, L"▾", -1, &arrow, DT_SINGLELINE | DT_VCENTER | DT_CENTER);
    }
    SelectObject(dc, old);
    EndPaint(hwnd, &ps);
}

void closeBar() {
    if (g_bar.hwnd) DestroyWindow(g_bar.hwnd);
}

void reportError(HWND owner, const std::wstring& text) { MessageBoxW(owner, text.c_str(), L"rethink", MB_ICONINFORMATION); }

void runCommand(IDispatch* app, HWND owner, const std::wstring& command, const GanttHit& hit) {
    try {
        ganttCommand(app, command, hit);
    } catch (const winrt::hresult_error& e) {
        reportError(owner, e.message().c_str());
    } catch (const std::exception& e) {
        reportError(owner, widen(e.what()));
    }
}

// Schaltfläche i: ihr Menü unter ihr zeigen, die Wahl ausführen, die Leiste mit dem neuen Stand neu aufbauen
void openTool(int i) {
    if (i < 0 || i >= static_cast<int>(g_bar.tools.size())) return;
    HWND hwnd = g_bar.hwnd;
    POINT at{g_bar.rects[i].left, g_bar.rects[i].bottom};
    ClientToScreen(hwnd, &at);
    g_bar.busy = true;
    std::wstring command = trackMenu(g_bar.tools[i].menu, hwnd, at, TPM_LEFTALIGN | TPM_TOPALIGN);
    if (command.ends_with(L"=?")) {
        static COLORREF custom[16] = {};
        CHOOSECOLORW cc{sizeof(cc)};
        cc.hwndOwner = GetWindow(hwnd, GW_OWNER);
        cc.lpCustColors = custom;
        cc.rgbResult = g_bar.tools[i].swatch >= 0 ? static_cast<COLORREF>(g_bar.tools[i].swatch) : RGB(0, 0, 0);
        cc.Flags = CC_FULLOPEN | CC_RGBINIT;
        command = ChooseColorW(&cc) ? std::format(L"{}#{:02X}{:02X}{:02X}", command.substr(0, command.size() - 1), GetRValue(cc.rgbResult),
                                                  GetGValue(cc.rgbResult), GetBValue(cc.rgbResult))
                                    : L"";
    }
    g_bar.busy = false;
    if (command.empty() || !IsWindow(hwnd)) return;
    runCommand(g_bar.app.Get(), GetWindow(hwnd, GW_OWNER), command, g_bar.hit);
    if (!IsWindow(hwnd)) return;
    try {
        g_bar.tools = buildToolbar(ganttState(g_bar.app.Get(), g_bar.hit), g_bar.hit);
    } catch (...) {
        return closeBar();  // Element nicht mehr ausgewählt
    }
    const SIZE size = layoutBar();
    SetWindowPos(hwnd, nullptr, 0, 0, size.cx, size.cy, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
    InvalidateRect(hwnd, nullptr, TRUE);
}

int toolAt(LPARAM lParam) {
    const POINT p{static_cast<short>(LOWORD(lParam)), static_cast<short>(HIWORD(lParam))};
    for (size_t i = 0; i < g_bar.rects.size(); ++i)
        if (PtInRect(&g_bar.rects[i], p)) return static_cast<int>(i);
    return -1;
}

LRESULT CALLBACK barProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
        case WM_MOUSEACTIVATE: return MA_NOACTIVATE;
        case WM_MOUSEMOVE: {
            const int hot = toolAt(lParam);
            if (hot != g_bar.hot) {
                g_bar.hot = hot;
                InvalidateRect(hwnd, nullptr, FALSE);
                TRACKMOUSEEVENT track{sizeof(track), TME_LEAVE, hwnd, 0};
                TrackMouseEvent(&track);
            }
            return 0;
        }
        case WM_MOUSELEAVE:
            g_bar.hot = -1;
            InvalidateRect(hwnd, nullptr, FALSE);
            return 0;
        case WM_LBUTTONDOWN: openTool(toolAt(lParam)); return 0;
        case WM_TIMER: {
            if (g_bar.busy) return 0;
            // schließen bei Esc, bei einem Klick daneben oder wenn PowerPoint nicht mehr vorn ist
            POINT cursor{};
            GetCursorPos(&cursor);
            RECT window;
            GetWindowRect(hwnd, &window);
            const bool pressed = (GetAsyncKeyState(VK_LBUTTON) | GetAsyncKeyState(VK_RBUTTON) | GetAsyncKeyState(VK_MBUTTON)) & 0x8000;
            DWORD process = 0;
            GetWindowThreadProcessId(GetForegroundWindow(), &process);
            if ((GetAsyncKeyState(VK_ESCAPE) & 0x8001) || (pressed && !PtInRect(&window, cursor)) || process != GetCurrentProcessId()) closeBar();
            return 0;
        }
        case WM_PAINT: paintBar(hwnd); return 0;
        case WM_DESTROY:
            KillTimer(hwnd, kBarTimer);
            g_bar.hwnd = nullptr, g_bar.app.Reset(), g_bar.tools.clear(), g_bar.hot = -1, g_bar.busy = false;
            return 0;
    }
    return DefWindowProcW(hwnd, message, wParam, lParam);
}

// Leiste über dem Mauspunkt zeigen (ersetzt eine offene); keine Werkzeuge für die Stelle: keine Leiste
void showBar(IDispatch* app, HWND owner, POINT screen, const GanttHit& hit, const GanttState& state) {
    closeBar();
    std::vector<Tool> tools = buildToolbar(state, hit);
    if (tools.empty()) return;
    static const bool registered = [] {
        WNDCLASSW wc{};
        wc.lpfnWndProc = barProc;
        wc.hInstance = GetModuleHandleW(nullptr);
        wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kBarClass;
        return RegisterClassW(&wc) != 0;
    }();
    if (!registered) return;
    if (!g_bar.font) {
        NONCLIENTMETRICSW metrics{sizeof(metrics)};
        SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, sizeof(metrics), &metrics, 0);
        g_bar.font = CreateFontIndirectW(&metrics.lfMenuFont);
    }
    g_bar.app = app, g_bar.hit = hit, g_bar.tools = std::move(tools);
    const SIZE size = layoutBar();
    MONITORINFO monitor{sizeof(monitor)};
    GetMonitorInfoW(MonitorFromPoint(screen, MONITOR_DEFAULTTONEAREST), &monitor);
    const RECT& work = monitor.rcWork;
    const int x = std::clamp<int>(screen.x - px(12), work.left, std::max<int>(work.left, work.right - size.cx));
    const int y = screen.y - size.cy - px(12) >= work.top ? screen.y - size.cy - px(12) : screen.y + px(24);
    g_bar.hwnd = CreateWindowExW(WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_NOACTIVATE, kBarClass, L"", WS_POPUP | WS_BORDER, x, y, size.cx, size.cy,
                                 owner, nullptr, GetModuleHandleW(nullptr), nullptr);
    if (!g_bar.hwnd) return;
    SetTimer(g_bar.hwnd, kBarTimer, 40, nullptr);
    ShowWindow(g_bar.hwnd, SW_SHOWNOACTIVATE);
    UpdateWindow(g_bar.hwnd);
}

}  // namespace

void showGanttMenu(IDispatch* app, HWND owner, POINT screen, const GanttHit& hit) {
    const GanttState state = ganttState(app, hit);
    showBar(app, owner, screen, hit, state);
    g_bar.busy = true;
    const std::wstring command = trackMenu(buildMenu(state, hit), owner, screen, TPM_RIGHTBUTTON);
    g_bar.busy = false;
    if (!command.empty()) {
        closeBar();  // ein Befehl aus dem Menü schließt auch die Leiste
        runCommand(app, owner, command, hit);
        return;
    }
    // Klick auf die Leiste, während das Menü offen war: das Menü schließt, die Schaltfläche öffnet sich gleich
    if (g_bar.hwnd && (GetAsyncKeyState(VK_LBUTTON) & 0x8000)) {
        POINT cursor{};
        GetCursorPos(&cursor);
        ScreenToClient(g_bar.hwnd, &cursor);
        openTool(toolAt(MAKELPARAM(cursor.x, cursor.y)));
    }
}

HWND ganttToolbarWindow() { return g_bar.hwnd; }

std::wstring ganttMenuText(IDispatch* app, const GanttHit& hit) {
    std::wstring text;
    toText(buildMenu(ganttState(app, hit), hit), L"", text);
    return text;
}

std::wstring ganttToolbarText(IDispatch* app, const GanttHit& hit) {
    std::wstring text;
    for (const Tool& t : buildToolbar(ganttState(app, hit), hit)) {
        text += L"[" + t.label + (t.swatch >= 0 ? std::format(L" #{:02X}{:02X}{:02X}", GetRValue(t.swatch), GetGValue(t.swatch), GetBValue(t.swatch)) : L"") +
                L"]\n";
        toText(t.menu, t.label, text);
    }
    return text;
}

}  // namespace rethink::addin
