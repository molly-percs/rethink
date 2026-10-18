#include "dates.h"

#include <algorithm>
#include <cmath>
#include <cwctype>
#include <format>

namespace rethink::engine {
namespace {

bool isLeap(int y) { return (y % 4 == 0 && y % 100 != 0) || y % 400 == 0; }

int daysInMonth(int y, int m) {
    static constexpr int kDays[] = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
    return m == 2 && isLeap(y) ? 29 : kDays[m - 1];
}

long long floorDay(double t) { return static_cast<long long>(std::floor(t + 1e-9)); }

// Liest genau n Ziffern ab pos; -1 bei Fehler.
int digits(std::wstring_view s, size_t& pos, size_t n) {
    if (pos + n > s.size()) return -1;
    int value = 0;
    for (size_t i = 0; i < n; ++i) {
        const wchar_t c = s[pos + i];
        if (c < L'0' || c > L'9') return -1;
        value = value * 10 + (c - L'0');
    }
    pos += n;
    return value;
}

struct Names {
    const wchar_t* months[12];
    const wchar_t* monthsShort[12];
    const wchar_t* weekdays[7];
    const wchar_t* weekdaysShort[7];
    const wchar_t* week;
};

const Names& names(const std::wstring& language) {
    static const Names de{
        {L"Januar", L"Februar", L"März", L"April", L"Mai", L"Juni", L"Juli", L"August", L"September", L"Oktober", L"November", L"Dezember"},
        {L"Jan", L"Feb", L"Mär", L"Apr", L"Mai", L"Jun", L"Jul", L"Aug", L"Sep", L"Okt", L"Nov", L"Dez"},
        {L"Montag", L"Dienstag", L"Mittwoch", L"Donnerstag", L"Freitag", L"Samstag", L"Sonntag"},
        {L"Mo", L"Di", L"Mi", L"Do", L"Fr", L"Sa", L"So"},
        L"KW"};
    static const Names en{
        {L"January", L"February", L"March", L"April", L"May", L"June", L"July", L"August", L"September", L"October", L"November", L"December"},
        {L"Jan", L"Feb", L"Mar", L"Apr", L"May", L"Jun", L"Jul", L"Aug", L"Sep", L"Oct", L"Nov", L"Dec"},
        {L"Monday", L"Tuesday", L"Wednesday", L"Thursday", L"Friday", L"Saturday", L"Sunday"},
        {L"Mon", L"Tue", L"Wed", L"Thu", L"Fri", L"Sat", L"Sun"},
        L"Week"};
    return language == L"en" ? en : de;
}

std::wstring roman(int n, bool upper) {
    static const wchar_t* kRoman[] = {L"i", L"ii", L"iii", L"iv"};
    std::wstring s = n >= 1 && n <= 4 ? kRoman[n - 1] : std::to_wstring(n);
    if (upper)
        for (auto& c : s) c = static_cast<wchar_t>(std::towupper(c));
    return s;
}

}  // namespace

long long daysFromCivil(int year, int month, int day) {
    year -= month <= 2;
    const long long era = (year >= 0 ? year : year - 399) / 400;
    const unsigned yoe = static_cast<unsigned>(year - era * 400);
    const unsigned doy = (153 * (month > 2 ? month - 3 : month + 9) + 2) / 5 + day - 1;
    const unsigned doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    return era * 146097 + static_cast<long long>(doe) - 719468;
}

CivilDate civilFromDays(long long z) {
    z += 719468;
    const long long era = (z >= 0 ? z : z - 146096) / 146097;
    const unsigned doe = static_cast<unsigned>(z - era * 146097);
    const unsigned yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    const long long y = static_cast<long long>(yoe) + era * 400;
    const unsigned doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    const unsigned mp = (5 * doy + 2) / 153;
    const unsigned d = doy - (153 * mp + 2) / 5 + 1;
    const unsigned m = mp < 10 ? mp + 3 : mp - 9;
    return {static_cast<int>(y + (m <= 2)), static_cast<int>(m), static_cast<int>(d)};
}

std::optional<double> parseIsoDate(std::wstring_view s) {
    if (!s.empty() && (s.back() == L'Z' || s.back() == L'z')) s.remove_suffix(1);
    size_t pos = 0;
    const int year = digits(s, pos, 4);
    if (year < 0) return std::nullopt;
    int month = 1, day = 1, hour = 0, minute = 0, second = 0;
    const bool extended = pos < s.size() && s[pos] == L'-';
    auto timePart = [&](bool colons) -> bool {
        if (pos == s.size()) return true;
        if (s[pos] != L'T' && s[pos] != L't') return false;
        ++pos;
        if ((hour = digits(s, pos, 2)) < 0) return false;
        if (pos == s.size()) return true;
        if (colons && s[pos++] != L':') return false;
        if ((minute = digits(s, pos, 2)) < 0) return false;
        if (pos == s.size()) return true;
        if (colons && s[pos++] != L':') return false;
        return (second = digits(s, pos, 2)) >= 0;
    };
    if (extended) {
        ++pos;
        if ((month = digits(s, pos, 2)) < 0) return std::nullopt;
        if (pos < s.size() && s[pos] == L'-') {
            ++pos;
            if ((day = digits(s, pos, 2)) < 0 || !timePart(true)) return std::nullopt;
        }
    } else if (pos < s.size()) {
        if ((month = digits(s, pos, 2)) < 0) return std::nullopt;
        if (pos < s.size() && ((day = digits(s, pos, 2)) < 0 || !timePart(false))) return std::nullopt;
    }
    if (pos != s.size() || month < 1 || month > 12 || day < 1 || day > daysInMonth(year, month) || hour > 23 || minute > 59 || second > 59)
        return std::nullopt;
    return static_cast<double>(daysFromCivil(year, month, day)) + (hour * 3600 + minute * 60 + second) / 86400.0;
}

std::wstring isoDate(double t) {
    long long day = floorDay(t);
    long long minutes = std::llround((t - static_cast<double>(day)) * 1440.0);
    if (minutes >= 1440) ++day, minutes = 0;
    if (minutes < 0) minutes = 0;
    const CivilDate c = civilFromDays(day);
    if (minutes == 0) return std::format(L"{:04}-{:02}-{:02}", c.year, c.month, c.day);
    return std::format(L"{:04}-{:02}-{:02}T{:02}:{:02}", c.year, c.month, c.day, minutes / 60, minutes % 60);
}

int isoWeekday(long long days) { return static_cast<int>(((days % 7) + 7 + 3) % 7) + 1; }

int isoWeek(long long days) {
    const long long thursday = days - (isoWeekday(days) - 1) + 3;
    const int year = civilFromDays(thursday).year;
    return static_cast<int>((thursday - daysFromCivil(year, 1, 1)) / 7 + 1);
}

bool isWeekend(long long days, int first, int last) {
    const int wd = isoWeekday(days);
    return first <= last ? wd >= first && wd <= last : wd >= first || wd <= last;
}

long long unitStart(double t, TimeUnit unit, int weekStart) {
    const long long d = floorDay(t);
    const CivilDate c = civilFromDays(d);
    switch (unit) {
        case TimeUnit::Day: return d;
        case TimeUnit::Week: return d - (isoWeekday(d) - weekStart + 7) % 7;
        case TimeUnit::Month: return daysFromCivil(c.year, c.month, 1);
        case TimeUnit::Quarter: return daysFromCivil(c.year, (c.month - 1) / 3 * 3 + 1, 1);
        case TimeUnit::Year: return daysFromCivil(c.year, 1, 1);
    }
    return d;
}

long long nextUnit(long long start, TimeUnit unit) {
    switch (unit) {
        case TimeUnit::Day: return start + 1;
        case TimeUnit::Week: return start + 7;
        default: return static_cast<long long>(std::llround(addUnits(static_cast<double>(start), unit, 1)));
    }
}

double snapToUnit(double t, TimeUnit unit, int weekStart) {
    const long long s = unitStart(t, unit, weekStart), n = nextUnit(s, unit);
    return t - static_cast<double>(s) <= static_cast<double>(n) - t ? static_cast<double>(s) : static_cast<double>(n);
}

double addUnits(double t, TimeUnit unit, int n) {
    if (unit == TimeUnit::Day) return t + n;
    if (unit == TimeUnit::Week) return t + 7.0 * n;
    const long long day = floorDay(t);
    const double fraction = t - static_cast<double>(day);
    const CivilDate c = civilFromDays(day);
    const int months = unit == TimeUnit::Month ? n : unit == TimeUnit::Quarter ? 3 * n : 12 * n;
    int total = c.year * 12 + (c.month - 1) + months;
    const int year = total >= 0 ? total / 12 : (total - 11) / 12;
    const int month = total - year * 12 + 1;
    const int d = std::min(c.day, daysInMonth(year, month));
    return static_cast<double>(daysFromCivil(year, month, d)) + fraction;
}

double unitDays(TimeUnit unit) {
    switch (unit) {
        case TimeUnit::Day: return 1;
        case TimeUnit::Week: return 7;
        case TimeUnit::Month: return 30.44;
        case TimeUnit::Quarter: return 91.31;
        case TimeUnit::Year: return 365.25;
    }
    return 1;
}

std::wstring formatDate(double t, std::wstring_view code, const std::wstring& language) {
    const long long day = floorDay(t);
    const CivilDate c = civilFromDays(day);
    const Names& n = names(language);
    std::wstring out;
    for (size_t i = 0; i < code.size();) {
        const wchar_t ch = code[i];
        if (ch == L'"') {
            const size_t close = code.find(L'"', i + 1);
            out += code.substr(i + 1, (close == std::wstring_view::npos ? code.size() : close) - i - 1);
            i = close == std::wstring_view::npos ? code.size() : close + 1;
            continue;
        }
        if (ch == L'\\') {
            if (i + 1 < code.size()) out += code[i + 1];
            i += 2;
            continue;
        }
        if (ch == L'\'') {
            out += L'\'';
            i += i + 1 < code.size() && code[i + 1] == L'\'' ? 2 : 1;
            continue;
        }
        if (ch == L'¶') {
            out += L'\n';
            ++i;
            continue;
        }
        const wchar_t lower = static_cast<wchar_t>(std::towlower(ch));
        if (lower != L'd' && lower != L'w' && lower != L'm' && lower != L'q' && lower != L'y') {
            out += ch;
            ++i;
            continue;
        }
        size_t run = 1;
        while (i + run < code.size() && std::towlower(code[i + run]) == lower) ++run;
        i += run;
        const int wd = isoWeekday(day) - 1;
        switch (lower) {
            case L'd':
                if (run == 1) out += std::to_wstring(c.day);
                else if (run == 2) out += std::format(L"{:02}", c.day);
                else if (run == 3) out += n.weekdaysShort[wd];
                else if (run == 4) out += n.weekdays[wd];
                else out += std::wstring(1, n.weekdays[wd][0]);
                break;
            case L'w': {
                const int w = isoWeek(day);
                if (run == 1) out += std::to_wstring(w);
                else if (run == 2) out += std::format(L"{:02}", w);
                else if (run == 3) out += std::format(L"{} {}", n.week, w);
                else out += std::format(L"{} {:02}", n.week, w);
                break;
            }
            case L'm':
                if (run == 1) out += std::to_wstring(c.month);
                else if (run == 2) out += std::format(L"{:02}", c.month);
                else if (run == 3) out += n.monthsShort[c.month - 1];
                else if (run == 4) out += n.months[c.month - 1];
                else out += std::wstring(1, n.months[c.month - 1][0]);
                break;
            case L'q': {
                const int q = (c.month - 1) / 3 + 1;
                out += run == 1 ? std::to_wstring(q) : roman(q, run >= 3);
                break;
            }
            case L'y':
                out += run <= 2 ? std::format(L"{:02}", ((c.year % 100) + 100) % 100) : std::format(L"{:04}", c.year);
                break;
        }
    }
    return out;
}

std::wstring unitName(TimeUnit unit, const std::wstring& language) {
    const bool en = language == L"en";
    switch (unit) {
        case TimeUnit::Day: return en ? L"Day" : L"Tag";
        case TimeUnit::Week: return en ? L"Week" : L"Woche";
        case TimeUnit::Month: return en ? L"Month" : L"Monat";
        case TimeUnit::Quarter: return en ? L"Quarter" : L"Quartal";
        case TimeUnit::Year: return en ? L"Year" : L"Jahr";
    }
    return {};
}

}  // namespace rethink::engine
