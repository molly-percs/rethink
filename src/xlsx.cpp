#include "xlsx.h"

#include "opc.h"
#include "xmldoc.h"

#include <chrono>
#include <cmath>
#include <cwctype>
#include <format>
#include <map>
#include <tuple>

namespace rethink {
namespace {

constexpr wchar_t kRelOfficeDocument[] = L"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
constexpr wchar_t kRelWorksheet[] = L"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
constexpr wchar_t kRelStyles[] = L"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
constexpr wchar_t kRelSharedStrings[] = L"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

constexpr int kFmtGeneral = 0, kFmtPercent = 9, kFmtDate = 14;

// Stilschlüssel: (Zahlenformat, Füllfarbe oder -1)
using StyleKey = std::tuple<int, long long>;

class Styles {
public:
    int index(int numFmt, const std::optional<std::uint32_t>& fill) {
        const StyleKey key{numFmt, fill ? static_cast<long long>(*fill) : -1};
        if (auto it = xfs_.find(key); it != xfs_.end()) return it->second;
        int fillId = 0;
        if (fill) {
            auto f = fills_.find(*fill);
            if (f == fills_.end()) f = fills_.emplace(*fill, static_cast<int>(fills_.size()) + 2).first;
            fillId = f->second;
        }
        order_.push_back({numFmt, fillId});
        return xfs_[key] = static_cast<int>(order_.size());  // 0 ist der Standardstil
    }

    std::wstring xml() const {
        std::wstring fills = L"<fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill>";
        std::vector<std::uint32_t> byId(fills_.size());
        for (auto& [rgb, id] : fills_) byId[id - 2] = rgb;
        for (auto rgb : byId)
            fills += std::format(L"<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF{:06X}\"/><bgColor indexed=\"64\"/></patternFill></fill>", rgb);
        std::wstring xfs = L"<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>";
        for (auto& [fmt, fill] : order_)
            xfs += std::format(L"<xf numFmtId=\"{}\" fontId=\"0\" fillId=\"{}\" borderId=\"0\" xfId=\"0\"{}{}/>", fmt, fill,
                               fmt ? L" applyNumberFormat=\"1\"" : L"", fill ? L" applyFill=\"1\"" : L"");
        return std::format(
            LR"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
            LR"(<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">)"
            LR"(<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>)"
            LR"(<fills count="{}">{}</fills>)"
            LR"(<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>)"
            LR"(<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>)"
            LR"(<cellXfs count="{}">{}</cellXfs>)"
            LR"(<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>)",
            fills_.size() + 2, fills, order_.size() + 1, xfs);
    }

private:
    std::map<StyleKey, int> xfs_;
    std::map<std::uint32_t, int> fills_;
    std::vector<std::pair<int, int>> order_;
};

std::wstring sheetXml(const std::vector<SheetRow>& rows, Styles& styles) {
    std::wstring xml =
        LR"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
        LR"(<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>)";
    for (size_t r = 0; r < rows.size(); ++r) {
        xml += std::format(L"<row r=\"{}\">", r + 1);
        for (size_t c = 0; c < rows[r].size(); ++c) {
            const SheetCell& cell = rows[r][c];
            const std::wstring ref = columnName(c) + std::to_wstring(r + 1);
            auto style = [&](int fmt) {
                const int s = (fmt == kFmtGeneral && !cell.fill) ? 0 : styles.index(fmt, cell.fill);
                return s ? std::format(L" s=\"{}\"", s) : std::wstring();
            };
            switch (cell.kind) {
                case SheetCell::Kind::Empty:
                    if (cell.fill) xml += std::format(L"<c r=\"{}\"{}/>", ref, style(kFmtGeneral));
                    break;
                case SheetCell::Kind::Text:
                    xml += std::format(L"<c r=\"{}\"{} t=\"inlineStr\"><is><t xml:space=\"preserve\">{}</t></is></c>", ref,
                                       style(kFmtGeneral), xmlEscapeText(cell.text));
                    break;
                case SheetCell::Kind::Number:
                    xml += std::format(L"<c r=\"{}\"{}><v>{}</v></c>", ref, style(kFmtGeneral), formatNumber(cell.value));
                    break;
                case SheetCell::Kind::Percent:
                    xml += std::format(L"<c r=\"{}\"{}><v>{}</v></c>", ref, style(kFmtPercent), formatNumber(cell.value));
                    break;
                case SheetCell::Kind::Date:
                    xml += std::format(L"<c r=\"{}\"{}><v>{}</v></c>", ref, style(kFmtDate), formatNumber(cell.value));
                    break;
            }
        }
        xml += L"</row>";
    }
    xml += L"</sheetData></worksheet>";
    return xml;
}

// ---- Lesen ----

std::wstring sNs(std::wstring_view local) { return std::wstring(L"s:") + std::wstring(local); }

bool isPercentFormat(int id, const std::wstring& code) {
    if (id == 9 || id == 10) return true;
    bool quoted = false;
    for (wchar_t c : code) {
        if (c == L'"') quoted = !quoted;
        if (!quoted && c == L'%') return true;
    }
    return false;
}

bool isDateFormat(int id, const std::wstring& code) {
    if ((id >= 14 && id <= 22) || (id >= 45 && id <= 47)) return true;
    if (code.empty()) return false;
    bool quoted = false, bracket = false;
    for (wchar_t c : code) {
        if (c == L'"') quoted = !quoted;
        if (!quoted && c == L'[') bracket = true;
        if (!quoted && c == L']') bracket = false;
        if (!quoted && !bracket && (c == L'y' || c == L'Y' || c == L'd' || c == L'D')) return true;
    }
    return false;
}

struct CellStyle {
    bool percent = false;
    bool date = false;
    std::optional<std::uint32_t> fill;
};

std::vector<CellStyle> readStyles(Package& pkg, IOpcPart* workbook) {
    std::vector<CellStyle> result;
    for (auto& rel : pkg.relationships(workbook)) {
        if (rel.type != kRelStyles || !rel.targetPart || !pkg.exists(rel.targetPart.Get())) continue;
        auto stream = pkg.stream(pkg.part(rel.targetPart.Get()).Get());
        XmlDoc doc = XmlDoc::load(stream.Get());
        std::map<int, std::wstring> codes;
        for (auto& f : doc.all(nullptr, L"/s:styleSheet/s:numFmts/s:numFmt"))
            codes[std::stoi(doc.attr(f.Get(), L"numFmtId"))] = doc.attr(f.Get(), L"formatCode");
        std::vector<std::optional<std::uint32_t>> fills;
        for (auto& f : doc.all(nullptr, L"/s:styleSheet/s:fills/s:fill")) {
            std::optional<std::uint32_t> rgb;
            Node pattern = doc.one(f.Get(), L"s:patternFill[@patternType='solid']/s:fgColor[@rgb]");
            if (pattern) rgb = static_cast<std::uint32_t>(std::stoul(doc.attr(pattern.Get(), L"rgb"), nullptr, 16) & 0xFFFFFF);
            fills.push_back(rgb);
        }
        for (auto& xf : doc.all(nullptr, L"/s:styleSheet/s:cellXfs/s:xf")) {
            CellStyle s;
            const std::wstring fmtText = doc.attr(xf.Get(), L"numFmtId");
            const int fmt = fmtText.empty() ? 0 : std::stoi(fmtText);
            const auto code = codes.contains(fmt) ? codes[fmt] : std::wstring();
            s.percent = isPercentFormat(fmt, code);
            s.date = !s.percent && isDateFormat(fmt, code);
            const std::wstring fillText = doc.attr(xf.Get(), L"fillId");
            const size_t fillId = fillText.empty() ? 0 : std::stoul(fillText);
            if (fillId < fills.size()) s.fill = fills[fillId];
            result.push_back(s);
        }
    }
    return result;
}

std::vector<std::wstring> readSharedStrings(Package& pkg, IOpcPart* workbook) {
    std::vector<std::wstring> strings;
    for (auto& rel : pkg.relationships(workbook)) {
        if (rel.type != kRelSharedStrings || !rel.targetPart || !pkg.exists(rel.targetPart.Get())) continue;
        auto stream = pkg.stream(pkg.part(rel.targetPart.Get()).Get());
        XmlDoc doc = XmlDoc::load(stream.Get());
        for (auto& si : doc.all(nullptr, L"/s:sst/s:si")) {
            std::wstring text;
            for (auto& t : doc.all(si.Get(), L"s:t | s:r/s:t")) text += nodeText(t.Get());
            strings.push_back(text);
        }
    }
    return strings;
}

// "B12" → Spalte 1, Zeile 11
std::pair<size_t, size_t> parseRef(const std::wstring& ref) {
    size_t col = 0, i = 0;
    while (i < ref.size() && std::iswalpha(ref[i])) col = col * 26 + (std::towupper(ref[i++]) - L'A' + 1);
    const size_t row = std::stoul(ref.substr(i));
    return {col - 1, row - 1};
}

}  // namespace

std::wstring columnName(size_t index) {
    std::wstring name;
    ++index;
    while (index > 0) {
        const size_t rest = (index - 1) % 26;
        name.insert(name.begin(), static_cast<wchar_t>(L'A' + rest));
        index = (index - 1) / 26;
    }
    return name;
}

std::vector<std::uint8_t> makeWorkbook(const std::vector<SheetRow>& rows, const std::wstring& sheetName) {
    Package pkg = Package::createEmpty();
    auto workbook = pkg.createPart(L"/xl/workbook.xml", L"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
    auto sheet = pkg.createPart(L"/xl/worksheets/sheet1.xml", L"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
    auto styles = pkg.createPart(L"/xl/styles.xml", L"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
    Styles styleTable;
    pkg.writeText(workbook.Get(),
                  std::format(LR"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
                              LR"(<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">)"
                              LR"(<sheets><sheet name="{}" sheetId="1" r:id="rId1"/></sheets></workbook>)",
                              xmlEscape(sheetName)));
    pkg.writeText(sheet.Get(), sheetXml(rows, styleTable));
    pkg.writeText(styles.Get(), styleTable.xml());
    pkg.addRelationship(nullptr, kRelOfficeDocument, pkg.uri(L"/xl/workbook.xml").Get());
    pkg.addRelationship(workbook.Get(), kRelWorksheet, pkg.uri(L"/xl/worksheets/sheet1.xml").Get(), L"rId1");
    pkg.addRelationship(workbook.Get(), kRelStyles, pkg.uri(L"/xl/styles.xml").Get(), L"rId2");
    return pkg.saveToBytes();
}

std::vector<SheetRow> readWorkbook(const std::vector<std::uint8_t>& bytes) {
    Package pkg = Package::openBytes(bytes);
    ComPtr<IOpcPart> workbook;
    for (auto& rel : pkg.relationships(nullptr))
        if (rel.type == kRelOfficeDocument && rel.targetPart) workbook = pkg.part(rel.targetPart.Get());
    if (!workbook) throw Error("Die Datei ist keine Excel-Arbeitsmappe");

    auto wbStream = pkg.stream(workbook.Get());
    XmlDoc wbDoc = XmlDoc::load(wbStream.Get());
    Node firstSheet = wbDoc.one(nullptr, L"/s:workbook/s:sheets/s:sheet");
    if (!firstSheet) throw Error("Die Arbeitsmappe enthält kein Tabellenblatt");
    const std::wstring sheetRel = wbDoc.attr(firstSheet.Get(), L"r:id");
    PartUri sheetUri;
    for (auto& rel : pkg.relationships(workbook.Get()))
        if (rel.id == sheetRel) sheetUri = rel.targetPart;
    if (!sheetUri) throw Error("Das erste Tabellenblatt fehlt");

    const auto styles = readStyles(pkg, workbook.Get());
    const auto strings = readSharedStrings(pkg, workbook.Get());
    auto sheetStream = pkg.stream(pkg.part(sheetUri.Get()).Get());
    XmlDoc sheet = XmlDoc::load(sheetStream.Get());

    std::vector<SheetRow> rows;
    for (auto& c : sheet.all(nullptr, L"/s:worksheet/s:sheetData/s:row/s:c")) {
        const auto [col, row] = parseRef(sheet.attr(c.Get(), L"r"));
        const std::wstring styleText = sheet.attr(c.Get(), L"s");
        const CellStyle style = !styleText.empty() && std::stoul(styleText) < styles.size() ? styles[std::stoul(styleText)] : CellStyle{};
        const std::wstring type = sheet.attr(c.Get(), L"t");
        Node v = sheet.one(c.Get(), L"s:v");

        SheetCell cell;
        cell.fill = style.fill;
        if (type == L"s" && v) {
            const size_t i = std::stoul(nodeText(v.Get()));
            cell.kind = SheetCell::Kind::Text;
            cell.text = i < strings.size() ? strings[i] : L"";
        } else if (type == L"inlineStr") {
            cell.kind = SheetCell::Kind::Text;
            for (auto& t : sheet.all(c.Get(), L"s:is/s:t | s:is/s:r/s:t")) cell.text += nodeText(t.Get());
        } else if ((type == L"str" || type == L"e") && v) {
            cell.kind = SheetCell::Kind::Text;
            cell.text = nodeText(v.Get());
        } else if (type == L"b" && v) {
            cell.kind = SheetCell::Kind::Text;
            cell.text = nodeText(v.Get()) == L"1" ? L"WAHR" : L"FALSCH";
        } else if (v && !nodeText(v.Get()).empty()) {
            cell.value = std::stod(nodeText(v.Get()));
            cell.kind = style.percent ? SheetCell::Kind::Percent : style.date ? SheetCell::Kind::Date : SheetCell::Kind::Number;
        }
        if (cell.kind == SheetCell::Kind::Empty && !cell.fill) continue;

        if (rows.size() <= row) rows.resize(row + 1);
        if (rows[row].size() <= col) rows[row].resize(col + 1);
        rows[row][col] = cell;
    }
    // Leerzellen am Zeilenende entfernen (eine Zeile ohne Inhalt bleibt als leere Zeile stehen)
    for (auto& r : rows)
        while (!r.empty() && r.back().kind == SheetCell::Kind::Empty) r.pop_back();
    while (!rows.empty() && rows.back().empty()) rows.pop_back();
    return rows;
}

double excelSerialFromIso(const std::wstring& iso) {
    std::wstring digits;
    for (wchar_t c : iso.substr(0, iso.find(L'T')))
        if (std::iswdigit(c)) digits += c;
    const int y = std::stoi(digits.substr(0, 4));
    const unsigned m = digits.size() >= 6 ? std::stoul(digits.substr(4, 2)) : 1;
    const unsigned d = digits.size() >= 8 ? std::stoul(digits.substr(6, 2)) : 1;
    using namespace std::chrono;
    const auto days = sys_days(year_month_day(year(y), month(m), day(d))) - sys_days(year_month_day(year(1899), month(12), day(30)));
    return static_cast<double>(days.count());
}

std::wstring isoFromExcelSerial(double serial) {
    using namespace std::chrono;
    const year_month_day ymd(sys_days(year_month_day(year(1899), month(12), day(30))) + days(static_cast<int>(std::floor(serial))));
    return std::format(L"{:04}-{:02}-{:02}", static_cast<int>(ymd.year()), static_cast<unsigned>(ymd.month()), static_cast<unsigned>(ymd.day()));
}

}  // namespace rethink
