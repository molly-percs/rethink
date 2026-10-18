#pragma once

// Kalenderrechnung für den Zeitplan (Gantt): reine Rechnung, ohne Zeitzonen.
// Ein Zeitpunkt ist eine Zahl von Tagen seit 1970-01-01 00:00 (Bruchteil = Uhrzeit).

#include "options.h"

#include <optional>
#include <string>
#include <string_view>

namespace rethink::engine {

struct CivilDate {
    int year = 1970, month = 1, day = 1;
};

long long daysFromCivil(int year, int month, int day);
CivilDate civilFromDays(long long days);

// ISO-8601 wie im .ppttc-Schema: YYYY, YYYYMM, YYYYMMDD[Thh[mm[ss]]] oder YYYY-MM[-DD[Thh[:mm[:ss]]]], optional Z.
// Fehlende Teile gelten als Anfang (Monat 1, Tag 1, 00:00).
std::optional<double> parseIsoDate(std::wstring_view text);
// "YYYY-MM-DD", mit Uhrzeit "YYYY-MM-DDThh:mm", wenn der Zeitpunkt nicht auf Mitternacht fällt.
std::wstring isoDate(double t);

int isoWeekday(long long days);           // 1 = Montag … 7 = Sonntag
int isoWeek(long long days);              // Kalenderwoche nach ISO 8601
bool isWeekend(long long days, int firstWeekendDay, int lastWeekendDay);

// Anfang der Einheit, in der t liegt, und Anfang der folgenden Einheit.
long long unitStart(double t, TimeUnit unit, int weekStart);
long long nextUnit(long long start, TimeUnit unit);
// Zeitpunkt auf die nächste Einheitengrenze runden (Tag, Wochenanfang, Monatsanfang …).
double snapToUnit(double t, TimeUnit unit, int weekStart);
// Um n Einheiten verschieben (Monate/Quartale/Jahre kalendergenau, Tag des Monats begrenzt).
double addUnits(double t, TimeUnit unit, int n);
// Ungefähre Länge einer Einheit in Tagen (für die Wahl der Skalen).
double unitDays(TimeUnit unit);

// Datum nach Formatcode (Groß-/Kleinschreibung egal): d dd ddd dddd ddddd, w ww www wwww, M MM MMM MMMM MMMMM,
// Q QQ QQQ, yy yyyy; "Text" und \x sind wörtlich, ' oder '' ergibt einen Apostroph, ¶ einen Zeilenumbruch.
std::wstring formatDate(double t, std::wstring_view code, const std::wstring& language);

std::wstring unitName(TimeUnit unit, const std::wstring& language);

}  // namespace rethink::engine
