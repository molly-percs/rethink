#include "cellinput.h"

#include "util.h"

#include <algorithm>
#include <cwctype>
#include <format>
#include <regex>

namespace rethink {
namespace {

std::wstring trim(const std::wstring& s) {
    const auto first = s.find_first_not_of(L" \t\r\n ");
    if (first == std::wstring::npos) return {};
    const auto last = s.find_last_not_of(L" \t\r\n ");
    return s.substr(first, last - first + 1);
}

// Zahl mit Dezimal- und optionalen Tausendertrennzeichen in "C"-Schreibweise bringen; leer bei Nicht-Zahl.
std::wstring normalizeNumber(std::wstring s, const std::wstring& decimalSeparator) {
    s.erase(std::remove_if(s.begin(), s.end(), [](wchar_t c) { return c == L' ' || c == L' ' || c == L' '; }), s.end());
    if (s.empty()) return {};
    const wchar_t dec = decimalSeparator.empty() ? L'.' : decimalSeparator[0];
    const wchar_t other = dec == L',' ? L'.' : L',';
    const bool hasDec = s.find(dec) != std::wstring::npos, hasOther = s.find(other) != std::wstring::npos;
    std::wstring out;
    if (hasDec && hasOther) {
        // Beide vorhanden: das letzte ist das Dezimaltrennzeichen
        const wchar_t last = s.find_last_of(L".,") != std::wstring::npos ? s[s.find_last_of(L".,")] : dec;
        for (wchar_t c : s) {
            if (c == last) out += L'.';
            else if (c != L'.' && c != L',') out += c;
        }
    } else if (hasOther) {
        // Nur das andere Zeichen: vor genau drei Ziffern als Tausendertrennzeichen, sonst als Dezimalpunkt
        const std::wregex thousands(std::wstring(LR"(^[-+]?\d{1,3}(\)") + other + LR"(\d{3})+$)");
        if (std::regex_match(s, thousands)) {
            for (wchar_t c : s)
                if (c != other) out += c;
        } else {
            for (wchar_t c : s) out += c == other ? L'.' : c;
        }
    } else {
        for (wchar_t c : s) out += c == dec ? L'.' : c;
    }
    static const std::wregex number(LR"(^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$)");
    return std::regex_match(out, number) ? out : std::wstring();
}

}  // namespace

Cell parseCellInput(const std::wstring& raw, const std::wstring& decimalSeparator) {
    const std::wstring text = trim(raw);
    Cell cell;
    if (text.empty()) return cell;
    if (text == L"e" || text == L"E") {
        cell.kind = CellKind::String;
        cell.text = L"e";
        return cell;
    }
    std::wsmatch m;
    static const std::wregex german(LR"(^(\d{1,2})\.(\d{1,2})\.(\d{4})$)");
    static const std::wregex iso(LR"(^(\d{4})-(\d{1,2})-(\d{1,2})$)");
    if (std::regex_match(text, m, german) || std::regex_match(text, m, iso)) {
        const bool isIso = text.find(L'-') != std::wstring::npos;
        const int y = std::stoi(m[isIso ? 1 : 3].str()), mo = std::stoi(m[2].str()), d = std::stoi(m[isIso ? 3 : 1].str());
        if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
            cell.kind = CellKind::Date;
            cell.text = std::format(L"{:04}-{:02}-{:02}", y, mo, d);
            return cell;
        }
    }
    if (text.back() == L'%') {
        const std::wstring n = normalizeNumber(text.substr(0, text.size() - 1), decimalSeparator);
        if (!n.empty()) {
            cell.kind = CellKind::Percentage;
            cell.number = std::stod(n);
            return cell;
        }
    }
    if (const std::wstring n = normalizeNumber(text, decimalSeparator); !n.empty()) {
        cell.kind = CellKind::Number;
        cell.number = std::stod(n);
        return cell;
    }
    cell.kind = CellKind::String;
    cell.text = text;
    return cell;
}

std::wstring displayCell(const Cell& cell, const std::wstring& decimalSeparator) {
    auto number = [&](double v) {
        std::wstring s = formatNumber(v);
        if (const auto dot = s.find(L'.'); dot != std::wstring::npos) s.replace(dot, 1, decimalSeparator);
        return s;
    };
    switch (cell.kind) {
        case CellKind::String: return cell.text;
        case CellKind::Number: return number(cell.number);
        case CellKind::Percentage: return number(cell.number) + L"%";
        case CellKind::Date:
            if (cell.text.size() >= 10) return cell.text.substr(8, 2) + L"." + cell.text.substr(5, 2) + L"." + cell.text.substr(0, 4);
            return cell.text;
        case CellKind::Empty: break;
    }
    return {};
}

}  // namespace rethink
