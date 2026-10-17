#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/Windows.Data.Json.h>

#include "ppttc.h"
#include "util.h"

// windows.h definiert GetObject als Makro; das kollidiert mit JsonValue::GetObject.
#ifdef GetObject
#undef GetObject
#endif

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
    const std::wstring type = o.HasKey(L"type") ? std::wstring(o.GetNamedString(L"type")) : L"waterfall";
    if (type == L"waterfall")
        options.type = engine::ChartType::Waterfall;
    else
        fail("chart", std::format("Diagrammtyp \"{}\" ist in rethink noch nicht umgesetzt (siehe docs/charting-ledger.md)", narrow(type)));

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
