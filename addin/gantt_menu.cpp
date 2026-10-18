#include <winrt/base.h>

#include "gantt_menu.h"

#include "engine/chart.h"
#include "engine/dates.h"

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

std::vector<Entry> buildMenu(const GanttState& s, const GanttHit& hit) {
    const GanttOptions& g = s.options;
    const bool onRow = hit.row >= gantt::kFirstActivityRow;
    std::vector<Entry> m;

    // Element unter der Maus
    if (s.targetKind == "bar") {
        std::vector<Entry> shapes;
        if (!s.targetIsItem) shapes.push_back(item(L"Wie Vorgabe (" + shapeLabel(s.barShape) + L")", L"rowShape=", s.rowShape.empty(), true));
        for (auto& sh : kBarShapes)
            shapes.push_back(item(shapeLabel(sh), L"rowShape=" + sh, (s.rowShape.empty() ? (s.targetIsItem ? L"bar" : L"") : s.rowShape) == sh, true));
        m.push_back(submenu(L"Form", shapes));
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
        std::vector<Entry> shapes;
        if (!s.targetIsItem) {
            shapes.push_back(item(L"Wie Vorgabe (" + shapeLabel(s.marker) + L")", L"rowMarker=", s.rowMarker.empty(), true));
            for (auto& sh : kMarkers) shapes.push_back(item(shapeLabel(sh), L"rowMarker=" + sh, s.rowMarker == sh, true));
        } else {
            for (auto& sh : kMarkers) shapes.push_back(item(shapeLabel(sh), L"milestoneShape=" + sh, s.marker == sh, true));
        }
        m.push_back(submenu(L"Form", shapes));
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
    std::vector<Entry> scales;
    for (TimeUnit u : {TimeUnit::Year, TimeUnit::Quarter, TimeUnit::Month, TimeUnit::Week, TimeUnit::Day}) {
        const bool shown = std::find(s.scales.begin(), s.scales.end(), u) != s.scales.end();
        scales.push_back(toggle(unitLabel(u), std::wstring(L"scale:") + unitKey(u), shown));
    }
    scales.push_back(separator());
    scales.push_back(item(L"Automatisch (nach Platz)", L"scales=auto", g.scales.empty()));
    m.push_back(submenu(L"Skalen", scales));

    std::vector<Entry> formats;
    for (TimeUnit u : s.scales) {
        const std::wstring& current = storedFormat(g, u);
        std::vector<Entry> list{item(L"Automatisch (nach Platz)", std::format(L"format:{}=", unitKey(u)), current.empty(), true)};
        for (const wchar_t* code : scaleFormats(u)) {
            std::wstring example = formatDate(s.today, code, s.language);
            for (auto& c : example)
                if (c == L'\n') c = L' ';
            list.push_back(item(example, std::format(L"format:{}={}", unitKey(u), code), current == code, true));
        }
        formats.push_back(submenu(unitLabel(u), list));
    }
    m.push_back(submenu(L"Format der Skalen", formats, !formats.empty()));

    // Linien und Wochenenden
    m.push_back(submenu(L"Trennlinien", {toggle(L"Senkrecht (an den Skaleneinheiten)", L"separators", s.separators),
                                         toggle(L"Waagerecht (zwischen den Zeilen)", L"rowSeparators", s.rowSeparators), separator(),
                                         toggle(L"Meilensteinlinien", L"milestoneLines", g.milestoneLines), toggle(L"Heute-Linie", L"today", g.todayLine)}));
    m.push_back(submenu(L"Wochenenden", {item(L"Normal", L"weekends=plain", g.weekends == WeekendMode::Plain, true),
                                         item(L"Leicht grau", L"weekends=shaded", g.weekends == WeekendMode::Shaded, true),
                                         item(L"Ausgeblendet", L"weekends=hidden", g.weekends == WeekendMode::Hidden, true)}));

    // Beschriftung
    std::vector<Entry> dateFormats;
    const std::wstring currentDate = g.dateFormat.empty() ? L"dd.MM." : g.dateFormat;
    for (const wchar_t* code : {L"dd.MM.", L"dd.MM.yy", L"dd.MM.yyyy", L"d. MMM", L"d. MMM yy", L"ddd dd.MM.", L"MMM yy"})
        dateFormats.push_back(item(formatDate(s.today, code, s.language), std::wstring(L"dateFormat=") + code, currentDate == code, true));
    m.push_back(submenu(L"Beschriftung",
                        {submenu(L"Datum an Balken", {item(L"Keins", L"barLabels=none", g.barLabels == BarLabel::None, true),
                                                      item(L"Zeitraum", L"barLabels=range", g.barLabels == BarLabel::Range, true),
                                                      item(L"Anfang", L"barLabels=start", g.barLabels == BarLabel::Start, true),
                                                      item(L"Ende", L"barLabels=end", g.barLabels == BarLabel::End, true)}),
                         toggle(L"Datum an Meilensteinen", L"milestoneLabels", g.milestoneLabels), submenu(L"Datumsformat", dateFormats), separator(),
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

HMENU toMenu(const std::vector<Entry>& entries, std::vector<std::wstring>& commands) {
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
            if (!e.sub.empty()) {
                info.fMask |= MIIM_SUBMENU;
                info.hSubMenu = toMenu(e.sub, commands);
            } else {
                commands.push_back(e.command);
                info.wID = static_cast<UINT>(commands.size());
            }
        }
        InsertMenuItemW(menu, position++, TRUE, &info);
    }
    return menu;
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

}  // namespace

void showGanttMenu(IDispatch* app, HWND owner, POINT screen, const GanttHit& hit) {
    std::vector<std::wstring> commands;
    const GanttState state = ganttState(app, hit);
    HMENU menu = toMenu(buildMenu(state, hit), commands);
    const UINT id = static_cast<UINT>(TrackPopupMenuEx(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON | TPM_NONOTIFY, screen.x, screen.y, owner, nullptr));
    DestroyMenu(menu);
    if (id == 0 || id > commands.size()) return;
    try {
        ganttCommand(app, commands[id - 1], hit);
    } catch (const winrt::hresult_error& e) {
        MessageBoxW(owner, e.message().c_str(), L"rethink", MB_ICONINFORMATION);
    } catch (const std::exception& e) {
        MessageBoxW(owner, widen(e.what()).c_str(), L"rethink", MB_ICONINFORMATION);
    }
}

std::wstring ganttMenuText(IDispatch* app, const GanttHit& hit) {
    std::wstring text;
    toText(buildMenu(ganttState(app, hit), hit), L"", text);
    return text;
}

}  // namespace rethink::addin
