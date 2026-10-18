#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/Windows.Data.Json.h>

#include "ppttc.h"
#include "util.h"

// windows.h definiert GetObject als Makro; das kollidiert mit JsonValue::GetObject.
#ifdef GetObject
#undef GetObject
#endif

#include <algorithm>
#include <format>
#include <fstream>
#include <iterator>
#include <regex>

using namespace winrt::Windows::Data::Json;

namespace rethink {
namespace {

// Datumsformate wie im Schema: YYYY, YYYYMM, YYYYMMDD[Thh[mm[ss]]] oder YYYY-MM[-DD[Thh[:mm[:ss]]]], optional Z.
const std::wregex kDate(
    LR"(^[0-9]{4}((0[1-9]|1[0-2])((0[1-9]|[1-2][0-9]|3[0-1])(T([0-1][0-9]|2[0-3])([0-5][0-9]([0-5][0-9])?)?)?)?|-(0[1-9]|1[0-2])(-(0[1-9]|[1-2][0-9]|3[0-1])(T([0-1][0-9]|2[0-3])(:[0-5][0-9](:[0-5][0-9])?)?)?)?)?Z?$)");
const std::wregex kHexColor(LR"(^#([0-9A-Fa-f]{6})$)");
const std::wregex kRgbColor(LR"(^rgb\(([0-9]+),([0-9]+),([0-9]+)\)$)");

[[noreturn]] void fail(const std::string& where, std::string_view message) {
    throw Error(std::format("{}: {}", where, message));
}

bool hasControlChars(std::wstring_view s) {
    for (wchar_t c : s)
        if (c < 0x20 || c == 0x7F) return true;
    return false;
}

std::wstring requireName(const JsonObject& o, const wchar_t* key, const std::string& where) {
    if (!o.HasKey(key)) fail(where, std::format("Schlüssel \"{}\" fehlt", narrow(key)));
    auto v = o.GetNamedValue(key);
    if (v.ValueType() != JsonValueType::String) fail(where, std::format("\"{}\" muss eine Zeichenkette sein", narrow(key)));
    std::wstring s(v.GetString());
    if (s.empty() || hasControlChars(s)) fail(where, std::format("\"{}\" darf nicht leer sein und keine Steuerzeichen enthalten", narrow(key)));
    return s;
}

std::uint32_t parseFill(const std::wstring& s, const std::string& where) {
    std::wsmatch m;
    if (std::regex_match(s, m, kHexColor)) return static_cast<std::uint32_t>(std::stoul(m[1].str(), nullptr, 16));
    if (std::regex_match(s, m, kRgbColor)) {
        std::uint32_t rgb = 0;
        for (int i = 1; i <= 3; ++i) {
            const auto part = m[i].str();
            const unsigned long value = part.size() > 6 ? 256 : std::stoul(part);
            if (value > 255) fail(where, "\"fill\": Farbanteile müssen zwischen 0 und 255 liegen");
            rgb = (rgb << 8) | static_cast<std::uint32_t>(value);
        }
        return rgb;
    }
    fail(where, "\"fill\" muss \"#rrggbb\" oder \"rgb(r,g,b)\" sein");
}

Cell parseCell(const IJsonValue& value, const std::string& where) {
    Cell cell;
    if (value.ValueType() == JsonValueType::Null) return cell;
    if (value.ValueType() != JsonValueType::Object) fail(where, "Zelle muss null oder ein Objekt sein");
    const JsonObject o = value.GetObject();

    int kinds = 0;
    if (o.HasKey(L"string")) {
        auto v = o.GetNamedValue(L"string");
        if (v.ValueType() != JsonValueType::String) fail(where, "\"string\" muss eine Zeichenkette sein");
        cell.kind = CellKind::String;
        cell.text = std::wstring(v.GetString());
        ++kinds;
    }
    if (o.HasKey(L"number")) {
        auto v = o.GetNamedValue(L"number");
        if (v.ValueType() != JsonValueType::Number) fail(where, "\"number\" muss eine Zahl sein");
        cell.kind = CellKind::Number;
        cell.number = v.GetNumber();
        ++kinds;
    }
    if (o.HasKey(L"percentage")) {
        auto v = o.GetNamedValue(L"percentage");
        if (v.ValueType() != JsonValueType::Number) fail(where, "\"percentage\" muss eine Zahl sein");
        cell.kind = CellKind::Percentage;
        cell.number = v.GetNumber();
        ++kinds;
    }
    if (o.HasKey(L"date")) {
        auto v = o.GetNamedValue(L"date");
        if (v.ValueType() != JsonValueType::String) fail(where, "\"date\" muss eine Zeichenkette sein");
        cell.kind = CellKind::Date;
        cell.text = std::wstring(v.GetString());
        if (!std::regex_match(cell.text, kDate)) fail(where, "\"date\" muss ISO 8601 entsprechen, z. B. \"2024-01-01\"");
        ++kinds;
    }
    if (kinds != 1) fail(where, "Zelle braucht genau einen der Schlüssel string, number, percentage oder date");

    if (o.HasKey(L"fill")) {
        auto v = o.GetNamedValue(L"fill");
        if (v.ValueType() != JsonValueType::String) fail(where, "\"fill\" muss eine Zeichenkette sein");
        cell.fill = parseFill(std::wstring(v.GetString()), where);
    }
    return cell;
}

DataItem parseDataItem(const IJsonValue& value, const std::string& where) {
    if (value.ValueType() != JsonValueType::Object) fail(where, "Eintrag muss ein Objekt sein");
    const JsonObject o = value.GetObject();
    DataItem item;
    item.name = requireName(o, L"name", where);
    if (!o.HasKey(L"table") || o.GetNamedValue(L"table").ValueType() != JsonValueType::Array)
        fail(where, "\"table\" muss ein Array von Zeilen sein");
    const JsonArray rows = o.GetNamedArray(L"table");
    for (uint32_t r = 0; r < rows.Size(); ++r) {
        const std::string rowWhere = std::format("{}.table[{}]", where, r);
        const IJsonValue rowValue = rows.GetAt(r);
        if (rowValue.ValueType() != JsonValueType::Array) fail(rowWhere, "Zeile muss ein Array sein");
        const JsonArray cells = rowValue.GetArray();
        Row row;
        for (uint32_t c = 0; c < cells.Size(); ++c) row.push_back(parseCell(cells.GetAt(c), std::format("{}[{}]", rowWhere, c)));
        item.table.push_back(std::move(row));
    }
    if (o.HasKey(L"chart")) {
        auto chart = o.GetNamedValue(L"chart");
        if (chart.ValueType() != JsonValueType::Object) fail(where, "\"chart\" muss ein Objekt sein");
        item.chartJson = std::wstring(chart.Stringify());
    }
    return item;
}

IJsonValue cellToJson(const Cell& cell) {
    if (cell.kind == CellKind::Empty) return JsonValue::CreateNullValue();
    JsonObject o;
    switch (cell.kind) {
        case CellKind::String: o.SetNamedValue(L"string", JsonValue::CreateStringValue(cell.text)); break;
        case CellKind::Date: o.SetNamedValue(L"date", JsonValue::CreateStringValue(cell.text)); break;
        case CellKind::Number: o.SetNamedValue(L"number", JsonValue::CreateNumberValue(cell.number)); break;
        case CellKind::Percentage: o.SetNamedValue(L"percentage", JsonValue::CreateNumberValue(cell.number)); break;
        case CellKind::Empty: break;
    }
    if (cell.fill) o.SetNamedValue(L"fill", JsonValue::CreateStringValue(std::format(L"#{:06X}", *cell.fill)));
    return o;
}

JsonObject parseObject(const std::wstring& json, const char* what) {
    JsonObject o{nullptr};
    if (!JsonObject::TryParse(json, o)) throw Error(std::format("{} ist kein gültiges JSON-Objekt", what));
    return o;
}

bool optionalBool(const JsonObject& o, const wchar_t* key, bool fallback) {
    if (!o.HasKey(key)) return fallback;
    auto v = o.GetNamedValue(key);
    if (v.ValueType() != JsonValueType::Boolean) fail("chart", std::format("\"{}\" muss true oder false sein", narrow(key)));
    return v.GetBoolean();
}

double optionalNumber(const JsonObject& o, const wchar_t* key, double fallback) {
    if (!o.HasKey(key)) return fallback;
    auto v = o.GetNamedValue(key);
    if (v.ValueType() != JsonValueType::Number) fail("chart", std::format("\"{}\" muss eine Zahl sein", narrow(key)));
    return v.GetNumber();
}

std::wstring optionalString(const JsonObject& o, const wchar_t* key, const std::wstring& fallback, const char* where = "chart") {
    if (!o.HasKey(key)) return fallback;
    auto v = o.GetNamedValue(key);
    if (v.ValueType() != JsonValueType::String) fail(where, std::format("\"{}\" muss eine Zeichenkette sein", narrow(key)));
    return std::wstring(v.GetString());
}

// Zeichenkette aus einer festen Auswahl; liefert den Index in `choices`.
template <size_t N>
int optionalChoice(const JsonObject& o, const wchar_t* key, const wchar_t* const (&choices)[N], int fallback, const char* where = "chart") {
    if (!o.HasKey(key)) return fallback;
    const std::wstring value = optionalString(o, key, L"", where);
    for (size_t i = 0; i < N; ++i)
        if (value == choices[i]) return static_cast<int>(i);
    std::string list;
    for (size_t i = 0; i < N; ++i) list += (i ? ", " : "") + std::string("\"") + narrow(choices[i]) + "\"";
    fail(where, std::format("\"{}\" muss einer der Werte {} sein", narrow(key), list));
}

std::vector<int> intList(const JsonObject& o, const wchar_t* key) {
    std::vector<int> result;
    if (!o.HasKey(key)) return result;
    auto v = o.GetNamedValue(key);
    if (v.ValueType() != JsonValueType::Array) fail("chart", std::format("\"{}\" muss ein Array von Zeilennummern sein", narrow(key)));
    for (auto const& r : v.GetArray()) {
        if (r.ValueType() != JsonValueType::Number) fail("chart", std::format("\"{}\" muss ein Array von Zeilennummern sein", narrow(key)));
        result.push_back(static_cast<int>(r.GetNumber()));
    }
    return result;
}

constexpr const wchar_t* kOrientations[] = {L"up", L"down", L"right", L"left"};
constexpr const wchar_t* kUnits[] = {L"day", L"week", L"month", L"quarter", L"year"};
constexpr const wchar_t* kWeekendModes[] = {L"plain", L"shaded", L"hidden"};
constexpr const wchar_t* kWeekends[] = {L"sat-sun", L"fri-sat", L"thu-fri", L"fri", L"sun"};
constexpr int kWeekendDays[][2] = {{6, 7}, {5, 6}, {4, 5}, {5, 5}, {7, 7}};
constexpr const wchar_t* kWeekdays[] = {L"monday", L"tuesday", L"wednesday", L"thursday", L"friday", L"saturday", L"sunday"};
constexpr const wchar_t* kBarLabels[] = {L"none", L"range", L"start", L"end"};
constexpr const wchar_t* kItemKinds[] = {L"bar", L"pentagon", L"chevron", L"milestone", L"bracket", L"shade"};

void checkDate(const std::wstring& value, const std::string& where, const char* key) {
    if (!value.empty() && !std::regex_match(value, kDate))
        fail(where, std::format("\"{}\" muss ein ISO-8601-Datum sein, z. B. \"2026-11-02\"", key));
}

void parseGantt(const JsonObject& o, engine::GanttOptions& g) {
    g.start = optionalString(o, L"start", g.start);
    g.end = optionalString(o, L"end", g.end);
    checkDate(g.start, "chart", "start");
    checkDate(g.end, "chart", "end");
    if (o.HasKey(L"scales")) {
        auto v = o.GetNamedValue(L"scales");
        if (v.ValueType() == JsonValueType::Array) {
            for (auto const& s : v.GetArray()) {
                if (s.ValueType() != JsonValueType::String) fail("chart", "\"scales\" muss \"auto\" oder ein Array wie [\"month\", \"week\"] sein");
                const std::wstring name(s.GetString());
                const auto it = std::find_if(std::begin(kUnits), std::end(kUnits), [&](const wchar_t* u) { return name == u; });
                if (it == std::end(kUnits)) fail("chart", std::format("unbekannte Skala \"{}\" (day, week, month, quarter, year)", narrow(name)));
                g.scales.push_back(static_cast<engine::TimeUnit>(it - std::begin(kUnits)));
            }
        } else if (v.ValueType() != JsonValueType::String || std::wstring(v.GetString()) != L"auto") {
            fail("chart", "\"scales\" muss \"auto\" oder ein Array wie [\"month\", \"week\"] sein");
        }
    }
    g.weekends = static_cast<engine::WeekendMode>(optionalChoice(o, L"weekends", kWeekendModes, static_cast<int>(g.weekends)));
    if (o.HasKey(L"weekend")) {
        const int w = optionalChoice(o, L"weekend", kWeekends, 0);
        g.weekendFirst = kWeekendDays[w][0], g.weekendLast = kWeekendDays[w][1];
    }
    g.weekStart = optionalChoice(o, L"weekStart", kWeekdays, g.weekStart - 1) + 1;
    g.separators = optionalBool(o, L"separators", g.separators);
    g.barLabels = static_cast<engine::BarLabel>(optionalChoice(o, L"barLabels", kBarLabels, static_cast<int>(g.barLabels)));
    g.milestoneLabels = optionalBool(o, L"milestoneLabels", g.milestoneLabels);
    g.milestoneLines = optionalBool(o, L"milestoneLines", g.milestoneLines);
    g.dateFormat = optionalString(o, L"dateFormat", g.dateFormat);
    if (o.HasKey(L"formats")) {
        auto v = o.GetNamedValue(L"formats");
        if (v.ValueType() != JsonValueType::Object) fail("chart", "\"formats\" muss ein Objekt wie {\"month\": \"MMM\"} sein");
        const JsonObject f = v.GetObject();
        g.dayFormat = optionalString(f, L"day", g.dayFormat, "chart.formats");
        g.weekFormat = optionalString(f, L"week", g.weekFormat, "chart.formats");
        g.monthFormat = optionalString(f, L"month", g.monthFormat, "chart.formats");
        g.quarterFormat = optionalString(f, L"quarter", g.quarterFormat, "chart.formats");
        g.yearFormat = optionalString(f, L"year", g.yearFormat, "chart.formats");
    }
    if (o.HasKey(L"responsible")) g.responsible = optionalBool(o, L"responsible", false);
    if (o.HasKey(L"remark")) g.remark = optionalBool(o, L"remark", false);
    g.headlines = optionalBool(o, L"headlines", g.headlines);
    g.rowSeparators = optionalBool(o, L"rowSeparators", g.rowSeparators);
    g.shadedRows = intList(o, L"shadedRows");
    if (o.HasKey(L"today")) {
        auto v = o.GetNamedValue(L"today");
        if (v.ValueType() == JsonValueType::Boolean) {
            g.todayLine = v.GetBoolean();
        } else {
            g.today = optionalString(o, L"today", L"");
            checkDate(g.today, "chart", "today");
            g.todayLine = !g.today.empty();
        }
    }
    if (o.HasKey(L"items")) {
        auto v = o.GetNamedValue(L"items");
        if (v.ValueType() != JsonValueType::Array) fail("chart", "\"items\" muss ein Array sein");
        const JsonArray items = v.GetArray();
        for (uint32_t i = 0; i < items.Size(); ++i) {
            const std::string where = std::format("chart.items[{}]", i);
            if (items.GetAt(i).ValueType() != JsonValueType::Object) fail(where, "Element muss ein Objekt sein");
            const JsonObject e = items.GetAt(i).GetObject();
            engine::GanttItem it;
            it.kind = static_cast<engine::GanttItem::Kind>(optionalChoice(e, L"kind", kItemKinds, 0, where.c_str()));
            if (e.HasKey(L"row")) {
                if (e.GetNamedValue(L"row").ValueType() != JsonValueType::Number) fail(where, "\"row\" muss eine Zeilennummer sein");
                it.row = static_cast<int>(e.GetNamedNumber(L"row"));
            }
            it.start = optionalString(e, L"start", L"", where.c_str());
            it.end = optionalString(e, L"end", L"", where.c_str());
            checkDate(it.start, where, "start");
            checkDate(it.end, where, "end");
            if (it.start.empty()) fail(where, "\"start\" fehlt");
            if (it.kind != engine::GanttItem::Kind::Milestone && it.end.empty()) fail(where, "\"end\" fehlt");
            it.text = optionalString(e, L"text", L"", where.c_str());
            if (e.HasKey(L"fill")) it.fill = parseFill(optionalString(e, L"fill", L"", where.c_str()), where);
            g.items.push_back(it);
        }
    }
}

}  // namespace

DataItem parseDataItemJson(const std::wstring& json) {
    JsonValue value{nullptr};
    if (!JsonValue::TryParse(json, value)) throw Error("RETHINK_MODEL ist kein gültiges JSON");
    return parseDataItem(value, "RETHINK_MODEL");
}

std::wstring serializeDataItem(const DataItem& item) {
    JsonObject o;
    o.SetNamedValue(L"name", JsonValue::CreateStringValue(item.name));
    if (!item.chartJson.empty()) o.SetNamedValue(L"chart", parseObject(item.chartJson, "\"chart\""));
    JsonArray rows;
    for (auto& row : item.table) {
        JsonArray cells;
        for (auto& cell : row) cells.Append(cellToJson(cell));
        rows.Append(cells);
    }
    o.SetNamedValue(L"table", rows);
    return std::wstring(o.Stringify());
}

std::wstring mergeChartJson(const std::wstring& stored, const std::wstring& override_) {
    if (stored.empty()) return override_;
    if (override_.empty()) return stored;
    JsonObject merged = parseObject(stored, "Gespeichertes \"chart\"");
    const JsonObject changes = parseObject(override_, "\"chart\"");
    for (auto const& entry : changes) merged.SetNamedValue(entry.Key(), entry.Value());
    return std::wstring(merged.Stringify());
}

engine::ChartOptions parseChartOptions(const std::wstring& json) {
    const JsonObject o = parseObject(json, "\"chart\"");
    engine::ChartOptions options;
    const std::wstring type = optionalString(o, L"type", L"waterfall");
    if (type == L"waterfall")
        options.type = engine::ChartType::Waterfall;
    else if (type == L"gantt")
        options.type = engine::ChartType::Gantt;
    else
        fail("chart", std::format("Diagrammtyp \"{}\" ist in rethink noch nicht umgesetzt (siehe docs/charting-ledger.md)", narrow(type)));
    options.orientation = static_cast<engine::Orientation>(optionalChoice(o, L"orientation", kOrientations, 0));
    if (options.type == engine::ChartType::Gantt) parseGantt(o, options.gantt);

    options.dataTable = optionalBool(o, L"datatable", options.dataTable);
    options.segmentLabels = optionalBool(o, L"labels", options.segmentLabels);
    options.categoryLabels = optionalBool(o, L"categories", options.categoryLabels);
    options.connectors = optionalBool(o, L"connectors", options.connectors);
    options.fontPt = optionalNumber(o, L"fontSize", options.fontPt);
    options.gap = optionalNumber(o, L"gap", options.gap);
    if (o.HasKey(L"decimals")) options.decimals = static_cast<int>(optionalNumber(o, L"decimals", 0));
    if (o.HasKey(L"tableRows")) {
        auto v = o.GetNamedValue(L"tableRows");
        if (v.ValueType() != JsonValueType::Array) fail("chart", "\"tableRows\" muss ein Array von Zeilennummern sein");
        for (auto const& r : v.GetArray()) {
            if (r.ValueType() != JsonValueType::Number) fail("chart", "\"tableRows\" muss ein Array von Zeilennummern sein");
            options.tableRows.push_back(static_cast<int>(r.GetNumber()));
        }
        if (!options.tableRows.empty()) options.dataTable = optionalBool(o, L"datatable", true);
    }
    if (options.fontPt < 4 || options.fontPt > 72) fail("chart", "\"fontSize\" muss zwischen 4 und 72 liegen");
    if (options.gap < 0 || options.gap > 5) fail("chart", "\"gap\" muss zwischen 0 und 5 liegen");
    return options;
}

std::wstring ganttEditsJson(const engine::GanttOptions& g) {
    JsonObject o;
    o.SetNamedValue(L"start", JsonValue::CreateStringValue(g.start));
    o.SetNamedValue(L"end", JsonValue::CreateStringValue(g.end));
    if (g.scales.empty()) {
        o.SetNamedValue(L"scales", JsonValue::CreateStringValue(L"auto"));
    } else {
        JsonArray scales;
        for (auto u : g.scales) scales.Append(JsonValue::CreateStringValue(kUnits[static_cast<int>(u)]));
        o.SetNamedValue(L"scales", scales);
    }
    o.SetNamedValue(L"weekends", JsonValue::CreateStringValue(kWeekendModes[static_cast<int>(g.weekends)]));
    for (size_t i = 0; i < std::size(kWeekends); ++i)
        if (kWeekendDays[i][0] == g.weekendFirst && kWeekendDays[i][1] == g.weekendLast) o.SetNamedValue(L"weekend", JsonValue::CreateStringValue(kWeekends[i]));
    o.SetNamedValue(L"weekStart", JsonValue::CreateStringValue(kWeekdays[std::clamp(g.weekStart, 1, 7) - 1]));
    o.SetNamedValue(L"separators", JsonValue::CreateBooleanValue(g.separators));
    o.SetNamedValue(L"barLabels", JsonValue::CreateStringValue(kBarLabels[static_cast<int>(g.barLabels)]));
    o.SetNamedValue(L"milestoneLabels", JsonValue::CreateBooleanValue(g.milestoneLabels));
    o.SetNamedValue(L"milestoneLines", JsonValue::CreateBooleanValue(g.milestoneLines));
    if (g.responsible) o.SetNamedValue(L"responsible", JsonValue::CreateBooleanValue(*g.responsible));
    if (g.remark) o.SetNamedValue(L"remark", JsonValue::CreateBooleanValue(*g.remark));
    o.SetNamedValue(L"headlines", JsonValue::CreateBooleanValue(g.headlines));
    o.SetNamedValue(L"rowSeparators", JsonValue::CreateBooleanValue(g.rowSeparators));
    if (!g.today.empty() && g.todayLine) o.SetNamedValue(L"today", JsonValue::CreateStringValue(g.today));
    else o.SetNamedValue(L"today", JsonValue::CreateBooleanValue(g.todayLine));
    JsonArray shaded;
    for (int r : g.shadedRows) shaded.Append(JsonValue::CreateNumberValue(r));
    o.SetNamedValue(L"shadedRows", shaded);
    JsonArray items;
    for (auto& it : g.items) {
        JsonObject e;
        e.SetNamedValue(L"kind", JsonValue::CreateStringValue(kItemKinds[static_cast<int>(it.kind)]));
        e.SetNamedValue(L"row", JsonValue::CreateNumberValue(it.row));
        e.SetNamedValue(L"start", JsonValue::CreateStringValue(it.start));
        if (!it.end.empty()) e.SetNamedValue(L"end", JsonValue::CreateStringValue(it.end));
        if (!it.text.empty()) e.SetNamedValue(L"text", JsonValue::CreateStringValue(it.text));
        if (it.fill) e.SetNamedValue(L"fill", JsonValue::CreateStringValue(std::format(L"#{:06X}", *it.fill)));
        items.Append(e);
    }
    o.SetNamedValue(L"items", items);
    return std::wstring(o.Stringify());
}

std::vector<SlideSpec> parsePpttc(const std::filesystem::path& file) {
    std::ifstream in(file, std::ios::binary);
    if (!in) throw Error("Datei nicht lesbar: " + narrow(file.wstring()));
    std::string bytes((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
    if (bytes.starts_with("\xEF\xBB\xBF")) bytes.erase(0, 3);

    JsonValue root{nullptr};
    if (!JsonValue::TryParse(widen(bytes), root)) throw Error("Keine gültige JSON-Datei: " + narrow(file.wstring()));
    if (root.ValueType() != JsonValueType::Array) throw Error("Die .ppttc-Datei muss ein JSON-Array von Folien enthalten");

    std::vector<SlideSpec> specs;
    const JsonArray entries = root.GetArray();
    for (uint32_t i = 0; i < entries.Size(); ++i) {
        const std::string where = std::format("[{}]", i);
        const IJsonValue entry = entries.GetAt(i);
        if (entry.ValueType() != JsonValueType::Object) fail(where, "Eintrag muss ein Objekt sein");
        const JsonObject o = entry.GetObject();
        SlideSpec spec;
        spec.templateRef = requireName(o, L"template", where);
        if (!o.HasKey(L"data") || o.GetNamedValue(L"data").ValueType() != JsonValueType::Array)
            fail(where, "\"data\" muss ein Array sein");
        const JsonArray data = o.GetNamedArray(L"data");
        for (uint32_t d = 0; d < data.Size(); ++d)
            spec.data.push_back(parseDataItem(data.GetAt(d), std::format("{}.data[{}]", where, d)));
        specs.push_back(std::move(spec));
    }
    return specs;
}

std::wstring cellText(const Cell& cell) {
    switch (cell.kind) {
        case CellKind::String:
        case CellKind::Date: return cell.text;
        case CellKind::Number: return formatNumber(cell.number);
        case CellKind::Percentage: return formatNumber(cell.number) + L"%";
        case CellKind::Empty: break;
    }
    return {};
}

}  // namespace rethink
