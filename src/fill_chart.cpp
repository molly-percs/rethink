#include "fill_chart.h"

#include "xlsx.h"
#include "xmldoc.h"

#include <algorithm>
#include <format>

namespace rethink {
namespace {

constexpr wchar_t kRelPackage[] = L"http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
constexpr wchar_t kXlsxContentType[] = L"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

struct Series {
    std::wstring name;
    std::vector<Cell> values;
    int colorIndex = 0;  // Zeilenindex - 1; Leerzeilen verbrauchen einen Index
};

struct ChartTable {
    std::vector<std::wstring> categories;
    std::vector<Series> series;
    size_t points = 0;
};

// Datumskategorien in der gröbsten Genauigkeit, die alle Werte trägt:
// lauter 1. Januar → "2024", lauter Monatserste → "2024-03", sonst das Datum wie angegeben.
std::vector<std::wstring> categoryLabels(const Row& head) {
    std::vector<std::wstring> labels;
    bool allDates = head.size() > 1, yearOnly = true, monthOnly = true;
    for (size_t c = 1; c < head.size(); ++c) {
        const Cell& cell = head[c];
        if (cell.kind != CellKind::Date) {
            allDates = false;
            break;
        }
        const std::wstring digits = [&] {
            std::wstring d;
            for (wchar_t ch : cell.text.substr(0, cell.text.find(L'T')))
                if (ch != L'-') d += ch;
            return d;
        }();
        const std::wstring month = digits.size() >= 6 ? digits.substr(4, 2) : L"01";
        const std::wstring day = digits.size() >= 8 ? digits.substr(6, 2) : L"01";
        const bool hasTime = cell.text.find(L'T') != std::wstring::npos;
        if (month != L"01" || day != L"01" || hasTime) yearOnly = false;
        if (day != L"01" || hasTime) monthOnly = false;
    }
    for (size_t c = 1; c < head.size(); ++c) {
        const Cell& cell = head[c];
        if (allDates && yearOnly)
            labels.push_back(cell.text.substr(0, 4));
        else if (allDates && monthOnly) {
            std::wstring digits;
            for (wchar_t ch : cell.text) if (ch != L'-') digits += ch;
            labels.push_back(digits.substr(0, 4) + L"-" + (digits.size() >= 6 ? digits.substr(4, 2) : L"01"));
        } else
            labels.push_back(cellText(cell));
    }
    return labels;
}

ChartTable interpret(const DataItem& item, std::vector<std::string>& warnings) {
    if (item.table.empty()) throw Error(std::format("Diagramm '{}': \"table\" ist leer", narrow(item.name)));
    ChartTable t;
    t.categories = categoryLabels(item.table.front());

    bool warnedText = false;
    for (size_t r = 1; r < item.table.size(); ++r) {
        const Row& row = item.table[r];
        if (row.empty()) continue;
        Series s;
        s.colorIndex = static_cast<int>(r - 1);
        s.name = cellText(row.front());
        s.values.assign(row.begin() + 1, row.end());
        for (auto& v : s.values) {
            if (v.kind == CellKind::String || v.kind == CellKind::Date) {
                if (!warnedText)
                    warnings.push_back(std::format("Diagramm '{}': Text- und Datumswerte in Datenzellen werden ignoriert", narrow(item.name)));
                warnedText = true;
                v = Cell{};
            }
        }
        t.points = std::max(t.points, s.values.size());
        t.series.push_back(std::move(s));
    }
    t.points = std::max({t.points, t.categories.size(), size_t{1}});
    if (t.series.empty()) throw Error(std::format("Diagramm '{}': keine Datenreihe in \"table\"", narrow(item.name)));
    return t;
}

double chartValue(const Cell& c) { return c.kind == CellKind::Percentage ? c.number / 100.0 : c.number; }

bool isNumeric(const Cell& c) { return c.kind == CellKind::Number || c.kind == CellKind::Percentage; }

std::wstring txXml(const std::wstring& name, size_t sheetRow) {
    return std::format(
        L"<c:tx xmlns:c=\"{}\"><c:strRef><c:f>Sheet1!$A${}</c:f><c:strCache><c:ptCount val=\"1\"/>"
        L"<c:pt idx=\"0\"><c:v>{}</c:v></c:pt></c:strCache></c:strRef></c:tx>",
        kNsC, sheetRow, xmlEscapeText(name));
}

std::wstring catXml(const std::vector<std::wstring>& categories) {
    std::wstring points;
    for (size_t i = 0; i < categories.size(); ++i)
        points += std::format(L"<c:pt idx=\"{}\"><c:v>{}</c:v></c:pt>", i, xmlEscapeText(categories[i]));
    return std::format(
        L"<c:cat xmlns:c=\"{}\"><c:strRef><c:f>Sheet1!$B$1:${}$1</c:f><c:strCache><c:ptCount val=\"{}\"/>{}"
        L"</c:strCache></c:strRef></c:cat>",
        kNsC, columnName(categories.size()), categories.size(), points);
}

std::wstring valXml(const Series& s, size_t points, size_t sheetRow) {
    size_t numeric = 0, percent = 0;
    for (auto& v : s.values) {
        if (isNumeric(v)) ++numeric;
        if (v.kind == CellKind::Percentage) ++percent;
    }
    const bool percentSeries = numeric > 0 && percent == numeric;
    std::wstring pts;
    for (size_t i = 0; i < s.values.size(); ++i) {
        const Cell& v = s.values[i];
        if (!isNumeric(v)) continue;
        const bool pointFormat = !percentSeries && v.kind == CellKind::Percentage;
        pts += std::format(L"<c:pt idx=\"{}\"{}><c:v>{}</c:v></c:pt>", i, pointFormat ? L" formatCode=\"0%\"" : L"",
                           formatNumber(chartValue(v)));
    }
    return std::format(
        L"<c:val xmlns:c=\"{}\"><c:numRef><c:f>Sheet1!$B${}:${}${}</c:f><c:numCache><c:formatCode>{}</c:formatCode>"
        L"<c:ptCount val=\"{}\"/>{}</c:numCache></c:numRef></c:val>",
        kNsC, sheetRow, columnName(points), sheetRow, percentSeries ? L"0%" : L"General", points, pts);
}

std::wstring dataPointXml(size_t index, std::uint32_t rgb) {
    const std::wstring hex = std::format(L"{:06X}", rgb);
    return std::format(
        L"<c:dPt xmlns:c=\"{}\" xmlns:a=\"{}\"><c:idx val=\"{}\"/><c:spPr><a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>"
        L"<a:ln><a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill></a:ln></c:spPr></c:dPt>",
        kNsC, kNsA, index, hex, hex);
}

// Baut eine Reihe aus einer Vorlagenreihe: Formatierung bleibt, Daten werden ersetzt.
Node buildSeries(const XmlDoc& doc, IXMLDOMNode* templateSeries, const Series& s, size_t order) {
    Node ser = cloneDeep(templateSeries);
    const size_t sheetRow = order + 2;

    if (Node idx = doc.one(ser.Get(), L"c:idx")) setAttr(idx.Get(), L"val", std::to_wstring(s.colorIndex));
    if (Node ord = doc.one(ser.Get(), L"c:order")) setAttr(ord.Get(), L"val", std::to_wstring(order));

    // Eindeutige IDs (c16:uniqueId) dürfen nicht dupliziert werden; Daten werden neu geschrieben.
    for (auto& n : doc.all(ser.Get(), L"c:extLst | c:tx | c:cat | c:val")) removeNode(n.Get());

    // Vorhandene Datenpunkt-Formate (z. B. Kreissegmente) bleiben, außer sie liegen außerhalb
    // der neuen Daten oder werden durch "fill" ersetzt.
    for (auto& point : doc.all(ser.Get(), L"c:dPt")) {
        const std::wstring idxText = doc.attr(doc.one(point.Get(), L"c:idx").Get(), L"val");
        const size_t idx = idxText.empty() ? SIZE_MAX : std::stoul(idxText);
        if (idx >= s.values.size() || s.values[idx].fill) removeNode(point.Get());
    }

    // Füllschema: Akzentfarben folgen dem Zeilenindex (Fläche/Linie und Marker).
    for (auto& clr : doc.all(ser.Get(), L"c:spPr//a:schemeClr[starts-with(@val,'accent')] | "
                                        L"c:marker/c:spPr//a:schemeClr[starts-with(@val,'accent')]"))
        setAttr(clr.Get(), L"val", std::format(L"accent{}", s.colorIndex % 6 + 1));

    Node tx = doc.fragment(txXml(s.name, sheetRow));
    insertBefore(ser.Get(), tx.Get(),
                 firstChildOf(ser.Get(), kNsC, {L"spPr", L"invertIfNegative", L"pictureOptions", L"marker", L"explosion",
                                                L"dLbls", L"trendline", L"errBars", L"smooth", L"shape", L"bubble3D"}).Get());

    for (size_t i = 0; i < s.values.size(); ++i) {
        if (!s.values[i].fill) continue;
        Node point = doc.fragment(dataPointXml(i, *s.values[i].fill));
        insertBefore(ser.Get(), point.Get(),
                     firstChildOf(ser.Get(), kNsC, {L"dLbls", L"trendline", L"errBars", L"smooth", L"shape", L"bubble3D"}).Get());
    }
    return ser;
}

void insertDataRefs(const XmlDoc& doc, IXMLDOMNode* ser, const ChartTable& t, const Series& s, size_t order) {
    if (!t.categories.empty()) {
        Node cat = doc.fragment(catXml(t.categories));
        insertBefore(ser, cat.Get(), firstChildOf(ser, kNsC, {L"smooth", L"shape", L"bubble3D"}).Get());
    }
    Node val = doc.fragment(valXml(s, t.points, order + 2));
    insertBefore(ser, val.Get(), firstChildOf(ser, kNsC, {L"smooth", L"shape", L"bubble3D"}).Get());
}

std::vector<SheetRow> sheetRows(const ChartTable& t) {
    std::vector<SheetRow> rows;
    SheetRow head(1);
    for (auto& c : t.categories) head.push_back({SheetCell::Kind::Text, c, 0});
    rows.push_back(std::move(head));
    for (auto& s : t.series) {
        SheetRow row;
        row.push_back({SheetCell::Kind::Text, s.name, 0});
        for (auto& v : s.values) {
            if (v.kind == CellKind::Number)
                row.push_back({SheetCell::Kind::Number, L"", v.number});
            else if (v.kind == CellKind::Percentage)
                row.push_back({SheetCell::Kind::Percent, L"", v.number / 100.0});
            else
                row.push_back({});
        }
        rows.push_back(std::move(row));
    }
    return rows;
}

void replaceWorkbook(Package& pkg, IOpcPart* chartPart, const XmlDoc& doc, const ChartTable& t) {
    const std::wstring partName = pkg.uniqueName(L"/ppt/embeddings/rethink_data1.xlsx");
    auto workbook = pkg.createPart(partName, kXlsxContentType);
    pkg.write(workbook.Get(), makeWorkbook(sheetRows(t)));
    PartUri workbookUri = pkg.uri(partName);

    Node external = doc.one(nullptr, L"/c:chartSpace/c:externalData");
    if (external) {
        const std::wstring id = doc.attr(external.Get(), L"r:id");
        pkg.deleteRelationship(chartPart, id);
        pkg.addRelationship(chartPart, kRelPackage, workbookUri.Get(), id);
        return;
    }
    const std::wstring id = pkg.addRelationship(chartPart, kRelPackage, workbookUri.Get());
    Node node = doc.fragment(std::format(L"<c:externalData xmlns:c=\"{}\" xmlns:r=\"{}\" r:id=\"{}\"><c:autoUpdate val=\"0\"/></c:externalData>",
                                         kNsC, kNsR, id));
    Node space = doc.one(nullptr, L"/c:chartSpace");
    insertBefore(space.Get(), node.Get(), firstChildOf(space.Get(), kNsC, {L"printSettings", L"userShapes", L"extLst"}).Get());
}

}  // namespace

void fillChart(Package& pkg, IOpcPart* chartPart, const DataItem& item, std::vector<std::string>& warnings) {
    auto content = pkg.stream(chartPart);
    XmlDoc doc = XmlDoc::load(content.Get());

    Node group = doc.one(nullptr, L"/c:chartSpace/c:chart/c:plotArea/*[c:ser]");
    if (!group) throw Error(std::format("Diagramm '{}' enthält keine Datenreihe, die als Vorlage dienen kann", narrow(item.name)));
    const std::wstring kind = localName(group.Get());
    if (kind == L"scatterChart" || kind == L"bubbleChart")
        throw Error(std::format("Diagramm '{}': Punkt- und Blasendiagramme unterstützt rethink in Phase 1 nicht", narrow(item.name)));

    const ChartTable t = interpret(item, warnings);
    if (kind == L"pieChart" || kind == L"doughnutChart") {
        if (t.series.size() > 1)
            warnings.push_back(std::format("Diagramm '{}': Kreisdiagramme zeigen nur die erste Reihe", narrow(item.name)));
    }

    const auto existing = doc.all(group.Get(), L"c:ser");
    Node anchor = nextSibling(existing.back().Get());
    std::vector<Node> built;
    for (size_t k = 0; k < t.series.size(); ++k) {
        IXMLDOMNode* source = existing[std::min(k, existing.size() - 1)].Get();
        Node ser = buildSeries(doc, source, t.series[k], k);
        insertDataRefs(doc, ser.Get(), t, t.series[k], k);
        built.push_back(ser);
    }
    for (auto& old : existing) removeNode(old.Get());
    for (auto& ser : built) insertBefore(group.Get(), ser.Get(), anchor.Get());

    // Ohne eigenen Titel würde PowerPoint bei genau einer Reihe deren Namen als Titel einblenden.
    if (!doc.one(nullptr, L"/c:chartSpace/c:chart/c:title")) {
        if (Node autoTitle = doc.one(nullptr, L"/c:chartSpace/c:chart/c:autoTitleDeleted")) setAttr(autoTitle.Get(), L"val", L"1");
    }

    replaceWorkbook(pkg, chartPart, doc, t);
    doc.save(content.Get());
}

}  // namespace rethink
