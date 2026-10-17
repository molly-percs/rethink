#pragma once

// Layout-Kern: reine Berechnung ohne COM und ohne XML.
// Ergebnis ist eine Liste von Grundformen in EMU, die die Ausgaben (OOXML, später COM) nur umsetzen.

#include "../ppttc.h"
#include "options.h"

#include <cstdint>
#include <optional>
#include <string>
#include <vector>

namespace rethink::engine {

using Emu = std::int64_t;
inline constexpr Emu kEmuPerPt = 12700;

struct Box {
    Emu x = 0, y = 0, w = 0, h = 0;
};

struct Color {
    enum class Kind { None, Scheme, Rgb } kind = Kind::None;
    std::wstring scheme;  // "accent1", "tx1", "bg1" …
    std::uint32_t rgb = 0;

    static Color none() { return {}; }
    static Color schemeColor(std::wstring name) {
        Color c;
        c.kind = Kind::Scheme;
        c.scheme = std::move(name);
        return c;
    }
    static Color rgbColor(std::uint32_t value) {
        Color c;
        c.kind = Kind::Rgb;
        c.rgb = value;
        return c;
    }
};

enum class Align { Left, Center, Right };

struct Prim {
    enum class Kind { Rect, Line, Text } kind = Kind::Rect;
    Emu x = 0, y = 0, w = 0, h = 0;  // Linie: von (x, y) nach (x + w, y + h)
    Color fill, stroke, textColor;
    double strokePt = 0.75;
    bool dashed = false;
    std::wstring text;
    double fontPt = 12;
    Align align = Align::Center;
    std::string role;   // eindeutiger Rollenname, z. B. "segment r1 c3"; Formname im Element
    std::string flags;  // Segmente: "pos", "neg" oder "e" (berechnet) – für Griffe im Add-in
    double value = 0;   // Segmente: dargestellter Wert (für den Maßstab beim Ziehen)
};

// Gleich gezeichnet? (Geometrie, Text, Farben, Kennung) – für Aktualisierungen nur geänderter Formen.
bool samePrim(const Prim& a, const Prim& b);

struct Drawing {
    Box bounds;
    std::vector<Prim> prims;
    std::vector<std::string> warnings;
};

// ---- Wasserfall ----------------------------------------------------------

struct WaterfallSegment {
    int row = 0;                // Index in "table"
    double value = 0;           // eingegebener oder berechneter Wert
    double lo = 0, hi = 0;      // Ausdehnung in Dateneinheiten
    bool computed = false;      // aus "e" berechnet
    bool percent = false;       // Eingabe war "percentage"
    std::optional<std::uint32_t> fill;
};

struct WaterfallColumn {
    std::vector<WaterfallSegment> segments;
    bool anchored = false;      // Säule mit "e": steht auf der Basislinie (Zwischensumme)
    double levelBefore = 0;
    double levelAfter = 0;      // Höhe des Verbinders zur nächsten Säule
};

// Wasserfall-Arithmetik nach dem dokumentierten Verhalten:
//  - Säulen ohne "e" schweben ab dem laufenden Stand und verschieben ihn um ihre Summe.
//  - Ein "e" ergänzt seine Säule bis zur Basislinie (Zwischensumme); der Stand bleibt.
//  - Ein "e" in der ersten Säule wird rückwärts berechnet (der Stand endet bei 0).
//  - Höchstens ein "e" pro Säule; gemischte Vorzeichen stapeln getrennt nach oben und unten.
std::vector<WaterfallColumn> computeWaterfall(const DataItem& item, const ChartOptions& options,
                                              std::vector<std::string>* warnings = nullptr);

Drawing layoutChart(const DataItem& item, const ChartOptions& options, Box bounds);

// Nachkommastellen der Beschriftungen (fest aus den Optionen oder automatisch 0–2);
// das Add-in nutzt sie als Schrittweite der ▲▼-Griffe (10^-Stellen).
int chartDecimals(const DataItem& item, const ChartOptions& options);

// Schrittweite beim Ändern eines einzelnen Werts (▲▼, Ziehen): so viele Nachkommastellen, wie der Wert selbst
// eingegeben wurde (12 → 1er-Schritte, 12,5 → 0,1); eine feste "decimals"-Option hat Vorrang. Höchstens 4.
int stepDecimals(const Cell& cell, const ChartOptions& options);

// Wert beim Ziehen: Startwert + Mausweg (in Pixeln, positiv = Richtung wachsender Werte) / Pixel je Einheit,
// gerundet auf die Schrittweite. Der Maßstab wird zu Beginn eingefroren.
double dragValue(double startValue, double pixels, double pixelsPerUnit, int decimals);

// Maßstab einer Zeichnung in EMU je Werteinheit, abgelesen am Segment mit dem größten Betrag
// (0, wenn alle Segmente 0 sind). Grundlage des eingefrorenen Maßstabs beim Ziehen.
double emuPerUnit(const Drawing& drawing);

// Zahl mit festen Nachkommastellen und gewähltem Dezimaltrennzeichen.
std::wstring formatValue(double value, int decimals, const std::wstring& decimalSeparator, bool percent);

// Ist die Zelle ein "e" (berechnetes Segment)?
bool isComputedCell(const Cell& cell);

}  // namespace rethink::engine
