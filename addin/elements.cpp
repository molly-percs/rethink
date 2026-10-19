#include <winrt/base.h>

#include "elements.h"

#include "calendar_window.h"
#include "category_dialog.h"
#include "datasheet.h"
#include "datasheet_window.h"
#include "dispatch.h"
#include "engine/chart.h"
#include "engine/dates.h"
#include "engine/gantt_edit.h"
#include "ppttc.h"
#include "render_element.h"
#include "xlsx.h"

#include <ocidl.h>
#include <shellapi.h>

#include <cmath>
#include <exception>
#include <filesystem>
#include <format>
#include <functional>
#include <limits>
#include <map>
#include <optional>
#include <set>
#include <sstream>
#include <thread>
#include <tuple>
#include <utility>

namespace rethink::addin {
namespace fs = std::filesystem;
using namespace engine;

namespace {

constexpr wchar_t kTitle[] = L"rethink";
constexpr wchar_t kPrefix[] = L"rethink:";
constexpr wchar_t kLayoutTag[] = L"RETHINK_LAYOUT";  // Layout-Rahmen und Gruppenrahmen beim letzten Zeichnen
constexpr wchar_t kStyleTag[] = L"RETHINK_STYLE";    // zuletzt gesetzte Farbe einer Form (eigene Änderungen erkennen)
constexpr wchar_t kSegmentTag[] = L"RETHINK_SEG";    // "pos", "neg", "e"; Zeitplan: "bar", "milestone", "bracket", "shade"
constexpr wchar_t kKindTag[] = L"RETHINK_KIND";      // "waterfall up|down|right|left" oder "gantt"
constexpr wchar_t kGanttTag[] = L"RETHINK_GANTT";    // Zeitplan: Zeitachse links;rechts;Skalen oben;unten (Punkt, beim Zeichnen)
constexpr long kMsoShapeRectangle = 1;
constexpr long kMsoTextOrientationHorizontal = 1;
constexpr long kMsoGroup = 6;
constexpr long kMsoAnchorMiddle = 3;
constexpr long kMsoLineDash = 4;
constexpr long kMsoLineSolid = 1;
constexpr long kMsoLineSysDot = 11;
constexpr long kPpSelectionShapes = 2;
constexpr long kPpSelectionText = 3;

// ---- Hilfen ----

// JSON (WinRT), OPC und MSXML laufen wie beim Erzeugen aus .ppttc in einem eigenen Thread mit eigenem Apartment.
template <class F>
auto onWorker(F&& body) -> decltype(body()) {
    using Result = decltype(body());
    std::optional<Result> result;
    std::exception_ptr error;
    std::thread worker([&] {
        try {
            winrt::init_apartment(winrt::apartment_type::single_threaded);
            try {
                result.emplace(body());
            } catch (...) {
                error = std::current_exception();
            }
            winrt::uninit_apartment();
        } catch (...) {
            error = std::current_exception();
        }
    });
    worker.join();
    if (error) std::rethrow_exception(error);
    return std::move(*result);
}

double pt(Emu v) { return static_cast<double>(v) / kEmuPerPt; }
Emu emu(double points) { return static_cast<Emu>(std::llround(points * kEmuPerPt)); }

std::wstring localeDecimalSeparator() {
    wchar_t buffer[8] = {};
    if (GetLocaleInfoEx(LOCALE_NAME_USER_DEFAULT, LOCALE_SDECIMAL, buffer, 8) > 0) return buffer;
    return L",";
}

long themeColor(const std::wstring& scheme) {
    if (scheme.starts_with(L"accent") && scheme.size() == 7) return 5 + (scheme[6] - L'1');  // msoThemeColorAccent1 = 5
    if (scheme == L"tx1") return 13;
    if (scheme == L"bg1") return 14;
    if (scheme == L"tx2") return 15;
    if (scheme == L"bg2") return 16;
    return 13;
}

std::wstring signature(const Color& c) {
    switch (c.kind) {
        case Color::Kind::Scheme: return c.brightness == 0 ? L"s:" + c.scheme : std::format(L"s:{}@{:.2f}", c.scheme, c.brightness);
        case Color::Kind::Rgb: return std::format(L"r:{:06X}", c.rgb);
        case Color::Kind::None: break;
    }
    return L"n";
}

void applyColor(IDispatch* colorFormat, const Color& c) {
    if (c.kind == Color::Kind::Scheme) {
        disp::put(colorFormat, L"ObjectThemeColor", themeColor(c.scheme));
        if (c.brightness != 0) disp::put(colorFormat, L"Brightness", c.brightness);  // Designfarbe heller/dunkler, folgt dem Master
    } else if (c.kind == Color::Kind::Rgb) {
        disp::put(colorFormat, L"RGB", static_cast<long>(((c.rgb >> 16) & 0xFF) | (c.rgb & 0xFF00) | ((c.rgb & 0xFF) << 16)));
    }
}

std::wstring tagValue(IDispatch* shape, const wchar_t* tag) {
    try {
        return disp::call(disp::object(shape, L"Tags").Get(), L"Item", tag).str();
    } catch (const Error&) {
        return {};
    }
}

void setTag(IDispatch* shape, const wchar_t* tag, const std::wstring& value) {
    disp::call(disp::object(shape, L"Tags").Get(), L"Add", std::wstring(tag), value);
}

void setGeometry(IDispatch* shape, const Prim& p) {
    const Emu x = p.w < 0 ? p.x + p.w : p.x, y = p.h < 0 ? p.y + p.h : p.y;
    disp::put(shape, L"Left", pt(x));
    disp::put(shape, L"Top", pt(y));
    disp::put(shape, L"Width", pt(std::abs(p.w)));
    disp::put(shape, L"Height", pt(std::abs(p.h)));
}

// msoAutoShapeType einer vordefinierten OOXML-Form
long autoShapeType(const std::string& preset) {
    if (preset == "diamond") return 4;     // msoShapeDiamond
    if (preset == "triangle") return 7;    // msoShapeIsoscelesTriangle
    if (preset == "homePlate") return 51;  // msoShapePentagon
    if (preset == "chevron") return 52;    // msoShapeChevron
    if (preset == "ellipse") return 9;     // msoShapeOval
    if (preset == "roundRect") return 5;   // msoShapeRoundedRectangle
    if (preset == "star5") return 92;      // msoShape5pointStar
    return kMsoShapeRectangle;
}

// Textrahmen ohne Ränder und ohne automatische Größe, Text mittig
void formatText(IDispatch* shape, const Prim& p) {
    auto frame = disp::object(shape, L"TextFrame");
    disp::put(frame.Get(), L"AutoSize", 0L);
    disp::put(frame.Get(), L"WordWrap", 0L);
    disp::put(frame.Get(), L"MarginLeft", 0.0);
    disp::put(frame.Get(), L"MarginRight", 0.0);
    disp::put(frame.Get(), L"MarginTop", 0.0);
    disp::put(frame.Get(), L"MarginBottom", 0.0);
    disp::put(frame.Get(), L"VerticalAnchor", kMsoAnchorMiddle);
    auto range = disp::object(frame.Get(), L"TextRange");
    disp::put(range.Get(), L"Text", p.text);
    auto font = disp::object(range.Get(), L"Font");
    disp::put(font.Get(), L"Size", p.fontPt);
    if (p.bold) disp::put(font.Get(), L"Bold", -1L);
    applyColor(disp::object(font.Get(), L"Color").Get(), p.textColor);
    disp::put(disp::object(range.Get(), L"ParagraphFormat").Get(), L"Alignment", p.align == Align::Left ? 1L : p.align == Align::Right ? 3L : 2L);
}

// Umriss eines Rechtecks/einer Form: Farbe, Stärke, gestrichelt; ohne Farbe unsichtbar
void applyOutline(IDispatch* shape, const Prim& p) {
    auto line = disp::object(shape, L"Line");
    if (p.stroke.kind == Color::Kind::None) {
        disp::put(line.Get(), L"Visible", 0L);
        return;
    }
    disp::put(line.Get(), L"Visible", -1L);
    applyColor(disp::object(line.Get(), L"ForeColor").Get(), p.stroke);
    disp::put(line.Get(), L"Weight", p.strokePt);
    disp::put(line.Get(), L"DashStyle", p.dashed ? kMsoLineDash : p.dotted ? kMsoLineSysDot : kMsoLineSolid);
}

// Signatur von Füllung und Umriss: nur bei Änderungen der Vorgabe neu setzen (eigene Formatierung bleibt)
std::wstring shapeSignature(const Prim& p) {
    std::wstring sig = signature(p.fill);
    if (p.stroke.kind != Color::Kind::None) sig += std::format(L"|{}|{:.2f}|{}", signature(p.stroke), p.strokePt, p.dashed ? L"d" : L"s");
    return sig;
}

// Eine Grundform über das Objektmodell zeichnen.
ComPtr<IDispatch> drawPrim(IDispatch* shapes, const Prim& p) {
    switch (p.kind) {
        case Prim::Kind::Rect:
        case Prim::Kind::Shape: {
            const long type = p.kind == Prim::Kind::Shape ? autoShapeType(p.shape) : kMsoShapeRectangle;
            auto shape = disp::object(shapes, L"AddShape", type, pt(p.x), pt(p.y), pt(p.w), pt(p.h));
            auto fill = disp::object(shape.Get(), L"Fill");
            disp::call(fill.Get(), L"Solid");
            applyColor(disp::object(fill.Get(), L"ForeColor").Get(), p.fill);
            if (p.transparency > 0) disp::put(fill.Get(), L"Transparency", p.transparency);
            applyOutline(shape.Get(), p);
            disp::put(disp::object(shape.Get(), L"Shadow").Get(), L"Visible", 0L);
            if (p.rotation != 0) disp::put(shape.Get(), L"Rotation", p.rotation);
            if (!p.text.empty()) formatText(shape.Get(), p);
            return shape;
        }
        case Prim::Kind::Line: {
            auto shape = disp::object(shapes, L"AddLine", pt(p.x), pt(p.y), pt(p.x + p.w), pt(p.y + p.h));
            auto line = disp::object(shape.Get(), L"Line");
            disp::put(line.Get(), L"Weight", p.strokePt);
            applyColor(disp::object(line.Get(), L"ForeColor").Get(), p.stroke);
            if (p.dashed || p.dotted) disp::put(line.Get(), L"DashStyle", p.dashed ? kMsoLineDash : kMsoLineSysDot);
            return shape;
        }
        case Prim::Kind::Text: {
            auto shape = disp::object(shapes, L"AddTextbox", kMsoTextOrientationHorizontal, pt(p.x), pt(p.y), pt(p.w), pt(p.h));
            formatText(shape.Get(), p);
            setGeometry(shape.Get(), p);
            return shape;
        }
    }
    throw Error("Unbekannte Grundform");
}

// Rolle, Farbsignatur und Segmentart an einer Form vermerken.
void markShape(IDispatch* shape, const Prim& p) {
    disp::put(shape, L"Name", widen(p.role));
    if (p.kind == Prim::Kind::Rect || p.kind == Prim::Kind::Shape) setTag(shape, kStyleTag, shapeSignature(p));
    if (p.kind == Prim::Kind::Text) setTag(shape, kStyleTag, signature(p.textColor));
    if (!p.flags.empty()) setTag(shape, kSegmentTag, widen(p.flags));
}

// Zeichnet alle Grundformen neu und gruppiert sie.
ComPtr<IDispatch> drawElement(IDispatch* shapes, const Drawing& d, const std::wstring& name) {
    static unsigned long serial = GetTickCount();
    const std::wstring stem = std::format(L"rethink tmp {:08X}", ++serial);
    std::vector<std::wstring> names;
    for (size_t i = 0; i < d.prims.size(); ++i) {
        auto shape = drawPrim(shapes, d.prims[i]);
        names.push_back(std::format(L"{} {}", stem, i + 1));
        disp::put(shape.Get(), L"Name", names.back());
    }

    // Shapes.Range(Array(...)).Group() – über eindeutige Zwischennamen, danach Rollennamen
    SAFEARRAY* array = SafeArrayCreateVector(VT_VARIANT, 0, static_cast<ULONG>(names.size()));
    if (!array) throw Error("Speicher für die Formenliste fehlt");
    for (LONG i = 0; i < static_cast<LONG>(names.size()); ++i) {
        VARIANT v;
        VariantInit(&v);
        v.vt = VT_BSTR;
        v.bstrVal = SysAllocString(names[i].c_str());
        SafeArrayPutElement(array, &i, &v);
        VariantClear(&v);
    }
    disp::Value list;
    list.raw().vt = VT_ARRAY | VT_VARIANT;
    list.raw().parray = array;
    auto range = disp::invoke(shapes, L"Range", DISPATCH_METHOD | DISPATCH_PROPERTYGET, &list, 1).object();
    auto group = disp::object(range.Get(), L"Group");
    auto items = disp::object(group.Get(), L"GroupItems");
    for (size_t i = 0; i < d.prims.size(); ++i) markShape(disp::object(items.Get(), L"Item", names[i]).Get(), d.prims[i]);

    disp::put(group.Get(), L"Name", name);
    disp::put(group.Get(), L"AlternativeText", std::wstring(kPrefix) + name);
    return group;
}

// Passt eine vorhandene Form an ihre Grundform an. Farben werden nur gesetzt, wenn sich die Vorgabe des Modells
// geändert hat – eigene Formatierungen bleiben so erhalten.
void updateShape(IDispatch* shape, const Prim& p) {
    if (p.kind == Prim::Kind::Rect || p.kind == Prim::Kind::Shape) {
        // Form gewechselt (z. B. Balken → Chevron, Meilenstein → Dreieck): Formtyp und Drehung anpassen
        const long type = p.kind == Prim::Kind::Shape ? autoShapeType(p.shape) : kMsoShapeRectangle;
        if (disp::call(shape, L"AutoShapeType").toLong() != type) disp::put(shape, L"AutoShapeType", type);
        if (std::fabs(disp::call(shape, L"Rotation").toDouble() - p.rotation) > 0.01) disp::put(shape, L"Rotation", p.rotation);
    }
    setGeometry(shape, p);
    if (p.kind == Prim::Kind::Rect || p.kind == Prim::Kind::Shape) {
        const auto sig = shapeSignature(p);
        if (tagValue(shape, kStyleTag) != sig) {
            applyColor(disp::object(disp::object(shape, L"Fill").Get(), L"ForeColor").Get(), p.fill);
            applyOutline(shape, p);
            setTag(shape, kStyleTag, sig);
        }
        if (tagValue(shape, kSegmentTag) != widen(p.flags)) setTag(shape, kSegmentTag, widen(p.flags));
        if (!p.text.empty()) {
            auto range = disp::object(disp::object(shape, L"TextFrame").Get(), L"TextRange");
            if (disp::call(range.Get(), L"Text").str() != p.text) disp::put(range.Get(), L"Text", p.text);
        }
    } else if (p.kind == Prim::Kind::Text) {
        auto range = disp::object(disp::object(shape, L"TextFrame").Get(), L"TextRange");
        if (disp::call(range.Get(), L"Text").str() != p.text) disp::put(range.Get(), L"Text", p.text);
        const auto sig = signature(p.textColor);
        if (tagValue(shape, kStyleTag) != sig) {
            applyColor(disp::object(disp::object(range.Get(), L"Font").Get(), L"Color").Get(), p.textColor);
            setTag(shape, kStyleTag, sig);
        }
    }
}

using RoleMap = std::map<std::wstring, ComPtr<IDispatch>>;

// Formen der Gruppe nach Rolle; nullopt, wenn sie nicht genau zu den Grundformen passen.
std::optional<RoleMap> shapesByRole(IDispatch* group, const Drawing& d) {
    auto items = disp::object(group, L"GroupItems");
    const long count = disp::call(items.Get(), L"Count").toLong();
    if (count != static_cast<long>(d.prims.size())) return std::nullopt;
    RoleMap byRole;
    for (long i = 1; i <= count; ++i) {
        auto shape = disp::object(items.Get(), L"Item", i);
        byRole[disp::call(shape.Get(), L"Name").str()] = shape;
    }
    for (auto& p : d.prims)
        if (!byRole.contains(widen(p.role))) return std::nullopt;
    return byRole;
}

// Passt die vorhandenen Formen an, wenn dieselben Rollen vorkommen.
bool updateInPlace(IDispatch* group, const Drawing& d) {
    auto byRole = shapesByRole(group, d);
    if (!byRole) return false;
    for (auto& p : d.prims) updateShape((*byRole)[widen(p.role)].Get(), p);
    return true;
}

struct Prepared {
    Drawing drawing;
    std::wstring model;
    std::wstring kind;  // RETHINK_KIND
};

std::wstring kindOf(const ChartOptions& options, const Drawing& d) {
    if (options.type == ChartType::Gantt) return L"gantt";
    static constexpr const wchar_t* kDirections[] = {L"up", L"down", L"right", L"left"};
    return std::wstring(L"waterfall ") + kDirections[static_cast<int>(d.valueDirection)];
}

// Optionen wie beim Zeichnen: aus dem "chart"-Objekt plus Umgebung (Dezimaltrennzeichen, Sprache, heute). Worker-Thread.
ChartOptions optionsOf(const DataItem& item) {
    ChartOptions options = parseChartOptions(item.chartJson);
    applyLocale(options);
    return options;
}

Drawing layoutOrThrow(const DataItem& item, const ChartOptions& options, Box bounds) {
    try {
        return layoutChart(item, options, bounds);
    } catch (const std::invalid_argument& e) {
        throw Error(std::format("Diagramm '{}': {}", narrow(item.name), e.what()));
    }
}

Prepared prepare(const DataItem& item, Box bounds) {
    return onWorker([&] {
        const ChartOptions options = optionsOf(item);
        Prepared p;
        p.drawing = layoutOrThrow(item, options, bounds);
        p.model = serializeDataItem(item);
        p.kind = kindOf(options, p.drawing);
        return p;
    });
}

struct Rect {
    double left = 0, top = 0, width = 0, height = 0;
};

Rect rectOf(IDispatch* shape) {
    return {disp::call(shape, L"Left").toDouble(), disp::call(shape, L"Top").toDouble(), disp::call(shape, L"Width").toDouble(),
            disp::call(shape, L"Height").toDouble()};
}

// Gespeicherter Layout-Rahmen (Punkt) und Gruppenrahmen beim letzten Zeichnen.
struct Layout {
    Rect box, group;
};

std::optional<Layout> storedLayout(IDispatch* group) {
    const std::wstring text = tagValue(group, kLayoutTag);
    if (text.empty()) return std::nullopt;
    std::wistringstream in(text);
    Layout l;
    wchar_t sep = 0;
    in >> l.box.left >> sep >> l.box.top >> sep >> l.box.width >> sep >> l.box.height >> sep >> l.group.left >> sep >> l.group.top >> sep >>
        l.group.width >> sep >> l.group.height;
    if (in.fail()) return std::nullopt;
    return l;
}

// Layout-Rahmen für ein erneutes Zeichnen: folgt Verschieben und Größenänderung der Gruppe
// (die Abstände zwischen Gruppenrahmen und Layout-Rahmen bleiben erhalten).
Box layoutBox(IDispatch* group) {
    const Rect now = rectOf(group);
    const auto stored = storedLayout(group);
    if (!stored) return {emu(now.left), emu(now.top), emu(now.width), emu(now.height)};
    const double insetL = stored->box.left - stored->group.left, insetT = stored->box.top - stored->group.top;
    const double insetR = (stored->group.left + stored->group.width) - (stored->box.left + stored->box.width);
    const double insetB = (stored->group.top + stored->group.height) - (stored->box.top + stored->box.height);
    return {emu(now.left + insetL), emu(now.top + insetT), emu(std::max(now.width - insetL - insetR, 10.0)),
            emu(std::max(now.height - insetT - insetB, 10.0))};
}

// Layout-Rahmen und aktuellen Gruppenrahmen vermerken (Grundlage für layoutBox und relayoutIfResized).
void storeLayout(IDispatch* group, const Box& b) {
    const Rect g = rectOf(group);
    setTag(group, kLayoutTag, std::format(L"{};{};{};{};{};{};{};{}", pt(b.x), pt(b.y), pt(b.w), pt(b.h), g.left, g.top, g.width, g.height));
}

// Zeitplan: Zeitachse und Skalenband (Folienpunkte) für die Griffe des Overlays vermerken.
void storeGanttGeometry(IDispatch* group, const Drawing& d) {
    if (!d.axis.valid) return;
    // Zeilen: Oberkante, Höhe und Anzahl der Aktivitätszeilen, linker Rand (für das Ziehen von Zeilen)
    double rowsTop = 0, rowH = 0;
    int rows = 0;
    for (auto& r : d.rows)
        if (r.row >= gantt::kFirstActivityRow) {
            if (rows++ == 0) rowsTop = pt(r.y);
            rowH = pt(r.h);
        }
    const std::wstring value = std::format(L"{};{};{};{};{};{};{};{}", pt(d.axis.x0), pt(d.axis.x0 + d.axis.w), pt(d.scaleTop), pt(d.scaleBottom),
                                           rowsTop, rowH, rows, pt(d.bounds.x));
    if (tagValue(group, kGanttTag) != value) setTag(group, kGanttTag, value);
}

// Zeichnet ein Element (neu oder an Ort und Stelle) und vermerkt Modell und Layout.
// Ist das Element (oder eine Form darin) gerade ausgewählt?
bool isSelected(IDispatch* slide, IDispatch* group) {
    try {
        auto app = disp::object(slide, L"Application");
        auto window = disp::object(app.Get(), L"ActiveWindow");
        auto selected = rethinkElementOf(disp::object(window.Get(), L"Selection").Get());
        return selected && disp::call(selected.Get(), L"Id").toLong() == disp::call(group, L"Id").toLong();
    } catch (const std::exception&) {
        return false;
    }
}

ComPtr<IDispatch> renderElement(IDispatch* slide, ComPtr<IDispatch> group, const DataItem& item, Box box) {
    const Prepared prepared = prepare(item, box);
    if (!group || !updateInPlace(group.Get(), prepared.drawing)) {
        long zorder = 0;
        bool reselect = false;
        if (group) {
            zorder = disp::call(group.Get(), L"ZOrderPosition").toLong();
            reselect = isSelected(slide, group.Get());  // ein neu gezeichnetes Element bleibt ausgewählt
            disp::call(group.Get(), L"Delete");
        }
        group = drawElement(disp::object(slide, L"Shapes").Get(), prepared.drawing, item.name);
        if (zorder > 0)
            for (int i = 0; i < 500 && disp::call(group.Get(), L"ZOrderPosition").toLong() > zorder; ++i)
                disp::call(group.Get(), L"ZOrder", 3L);  // msoSendBackward: alte Stapelposition
        if (reselect) {
            try {
                disp::call(group.Get(), L"Select");
            } catch (const std::exception&) {
            }
        }
    }
    setTag(group.Get(), kModelTag, prepared.model);
    if (tagValue(group.Get(), kKindTag) != prepared.kind) setTag(group.Get(), kKindTag, prepared.kind);
    storeLayout(group.Get(), prepared.drawing.bounds);
    storeGanttGeometry(group.Get(), prepared.drawing);
    return group;
}

ComPtr<IDispatch> activeWindow(IDispatch* app) {
    auto windows = disp::object(app, L"Windows");
    if (disp::call(windows.Get(), L"Count").toLong() < 1) throw Error("Es ist keine Präsentation geöffnet.");
    return disp::object(app, L"ActiveWindow");
}

ComPtr<IDispatch> currentSlide(IDispatch* app) {
    try {
        return disp::object(disp::object(activeWindow(app).Get(), L"View").Get(), L"Slide");
    } catch (const Error&) {
        throw Error("Bitte in der Normalansicht eine Folie auswählen.");
    }
}

ComPtr<IDispatch> selectedElement(IDispatch* app) {
    auto element = rethinkElementOf(disp::object(activeWindow(app).Get(), L"Selection").Get());
    if (!element) throw Error("Bitte zuerst ein rethink-Diagramm auswählen (Wasserfall einfügen oder per .ppttc erzeugen).");
    return element;
}

std::wstring freeName(IDispatch* slide, const std::wstring& stem) {
    auto shapes = disp::object(slide, L"Shapes");
    std::set<std::wstring> taken;
    const long count = disp::call(shapes.Get(), L"Count").toLong();
    for (long i = 1; i <= count; ++i) {
        auto shape = disp::object(shapes.Get(), L"Item", i);
        taken.insert(toLower(disp::call(shape.Get(), L"Name").str()));
        const auto alt = disp::call(shape.Get(), L"AlternativeText").str();
        if (alt.starts_with(kPrefix)) taken.insert(toLower(alt.substr(std::size(kPrefix) - 1)));
    }
    std::wstring name;
    for (int n = 1; name.empty() || taken.contains(toLower(name)); ++n) name = std::format(L"{}{}", stem, n);
    return name;
}

Cell textCell(const std::wstring& s) { return Cell{CellKind::String, s, 0, std::nullopt}; }
Cell numberCell(double v) { return Cell{CellKind::Number, L"", v, std::nullopt}; }
Cell dateCell(double day) { return Cell{CellKind::Date, isoDate(day), 0, std::nullopt}; }

const wchar_t* orientationName(Orientation o) {
    static constexpr const wchar_t* kNames[] = {L"up", L"down", L"right", L"left"};
    return kNames[static_cast<int>(o)];
}

// Aufbauend: Werte summieren sich zu e; abbauend (gespiegeltes Symbol im Original): die Summe steht links (e in der
// ersten Spalte, rückwärts berechnet) und wird schrittweise abgebaut.
DataItem sampleWaterfall(const std::wstring& name, bool buildDown, Orientation orientation) {
    DataItem item;
    item.name = name;
    item.chartJson = std::format(LR"({{"type":"waterfall","datatable":true,"orientation":"{}"}})", orientationName(orientation));
    if (buildDown) {
        item.table = {
            {Cell{}, textCell(L"Gesamt"), textCell(L"Personal"), textCell(L"Material"), textCell(L"Vertrieb"), textCell(L"Sonstiges")},
            {textCell(L"Wert"), textCell(L"e"), numberCell(-40), numberCell(-25), numberCell(-20), numberCell(-15)},
        };
    } else {
        item.table = {
            {Cell{}, textCell(L"Start"), textCell(L"Umsatz"), textCell(L"Kosten"), textCell(L"Zwischensumme"), textCell(L"Sonstiges"), textCell(L"Ende")},
            {textCell(L"Wert"), numberCell(100), numberCell(40), numberCell(-25), textCell(L"e"), numberCell(-10), textCell(L"e")},
        };
    }
    return item;
}

double todayDays() {
    SYSTEMTIME now{};
    GetLocalTime(&now);
    return static_cast<double>(daysFromCivil(now.wYear, now.wMonth, now.wDay));
}

// Neuer Zeitplan wie im Original: der aktuelle und die zwei folgenden Monate; Beispielaktivitäten darin.
DataItem sampleGantt(const std::wstring& name) {
    const double m0 = static_cast<double>(unitStart(todayDays(), TimeUnit::Month, 1));
    DataItem item;
    item.name = name;
    item.chartJson = std::format(LR"({{"type":"gantt","start":"{}","end":"{}"}})", isoDate(m0), isoDate(addUnits(m0, TimeUnit::Month, 3) - 1));
    item.table = {
        {textCell(L"Aktivität"), textCell(L"Start"), textCell(L"Ende"), textCell(L"Meilenstein"), textCell(L"Verantwortlich"), textCell(L"Bemerkung"),
         textCell(L"Kategorie"), textCell(L"Vorgänger"), textCell(L"Form"), textCell(L"Marker")},
        {textCell(L"Meilensteine"), Cell{}, Cell{}, dateCell(m0 + 84)},
        {textCell(L"Analyse"), dateCell(m0 + 2), dateCell(m0 + 16)},
        {textCell(L"Konzept"), dateCell(m0 + 14), dateCell(m0 + 35), dateCell(m0 + 34)},
        {textCell(L"Umsetzung"), dateCell(m0 + 31), dateCell(m0 + 70)},
        {textCell(L"Test"), dateCell(m0 + 63), dateCell(m0 + 80), dateCell(m0 + 79)},
    };
    return item;
}

ComPtr<IDispatch> findShape(IDispatch* slide, long id) {
    auto shapes = disp::object(slide, L"Shapes");
    const long count = disp::call(shapes.Get(), L"Count").toLong();
    for (long i = 1; i <= count; ++i) {
        auto shape = disp::object(shapes.Get(), L"Item", i);
        if (disp::call(shape.Get(), L"Id").toLong() == id) return shape;
    }
    return nullptr;
}

// ---- Datenblätter ----

struct Sheet {
    fs::path file;
    fs::file_time_type stamp{};
    ComPtr<IDispatch> slide;
    long shapeId = 0;
    std::wstring name;
    // Excel-Automation (sonst: Datei wird bei Strg+S neu eingelesen)
    ComPtr<IDispatch> workbook;
    ComPtr<IConnectionPoint> point;
    DWORD cookie = 0;
    bool dirty = false;
    ULONGLONG dirtyAt = 0;
    bool closePending = false;  // Nutzer schließt: im nächsten Takt übernehmen und dann selbst schließen
    bool closingByUs = false;
    bool closed = false;
};

std::map<std::wstring, Sheet> g_sheets;  // Schlüssel: Datei (klein)
ComPtr<IDispatch> g_excel;               // eigene Excel-Instanz für Datenblätter
UINT_PTR g_timer = 0;
bool g_busy = false;
DISPID g_sheetChange = 0x61C, g_beforeClose = 0x60A;  // aus der Typbibliothek überschrieben

bool isBusy(const std::exception& e) {
    const std::string text = e.what();
    return text.find("0x80010001") != std::string::npos || text.find("0x8001010A") != std::string::npos;
}

// Tabelle in das Element übernehmen (an Ort und Stelle).
void applyTable(Sheet& s, std::vector<Row> table) {
    auto group = findShape(s.slide.Get(), s.shapeId);
    if (!group) throw Error(std::format("Das Diagramm '{}' ist nicht mehr auf der Folie.", narrow(s.name)));
    const std::wstring stored = tagValue(group.Get(), kModelTag);
    DataItem item = onWorker([&] {
        DataItem parsed = parseDataItemJson(stored);
        parsed.table = std::move(table);
        return parsed;
    });
    auto result = renderElement(s.slide.Get(), group, item, layoutBox(group.Get()));
    s.shapeId = disp::call(result.Get(), L"Id").toLong();
}

// Aktuellen Inhalt der (offenen) Arbeitsmappe übernehmen, ohne sie zu speichern.
void flushWorkbook(Sheet& s) {
    const fs::path copy = fs::path(s.file).replace_extension(L".rethink-copy.xlsx");
    disp::call(s.workbook.Get(), L"SaveCopyAs", copy.wstring());
    const auto bytes = readFileShared(copy);
    std::error_code ec;
    fs::remove(copy, ec);
    applyTable(s, onWorker([&] { return tableFromDatasheet(bytes); }));
}

Sheet* sheetFor(const std::wstring& key) {
    auto it = g_sheets.find(key);
    return it == g_sheets.end() ? nullptr : &it->second;
}

class WorkbookSink final : public IDispatch {
public:
    WorkbookSink(std::wstring key, GUID iid) : key_(std::move(key)), iid_(iid) {}
    STDMETHODIMP QueryInterface(REFIID iid, void** out) override {
        if (!out) return E_POINTER;
        if (iid == IID_IUnknown || iid == IID_IDispatch || iid == iid_) {
            *out = static_cast<IDispatch*>(this);
            AddRef();
            return S_OK;
        }
        *out = nullptr;
        return E_NOINTERFACE;
    }
    STDMETHODIMP_(ULONG) AddRef() override { return InterlockedIncrement(&refs_); }
    STDMETHODIMP_(ULONG) Release() override {
        const ULONG n = InterlockedDecrement(&refs_);
        if (n == 0) delete this;
        return n;
    }
    STDMETHODIMP GetTypeInfoCount(UINT* c) override {
        if (c) *c = 0;
        return S_OK;
    }
    STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo** i) override {
        if (i) *i = nullptr;
        return E_NOTIMPL;
    }
    STDMETHODIMP GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*) override { return E_NOTIMPL; }
    STDMETHODIMP Invoke(DISPID id, REFIID, LCID, WORD, DISPPARAMS* params, VARIANT*, EXCEPINFO*, UINT*) override {
        Sheet* s = sheetFor(key_);
        if (!s) return S_OK;
        if (id == g_sheetChange) {
            s->dirty = true;  // nur vormerken; übernommen wird im Zeitgeber auf dem PowerPoint-Thread
            s->dirtyAt = GetTickCount64();
        } else if (id == g_beforeClose && !s->closingByUs) {
            // Während des Ereignisses nimmt Excel keine Aufrufe an. Daher das Schließen anhalten, im nächsten Takt
            // die letzten Änderungen (auch Zellfarben) übernehmen und die Mappe dann ohne Rückfrage selbst schließen.
            if (params && params->cArgs >= 1) {
                VARIANT& cancel = params->rgvarg[0];
                if (cancel.vt == (VT_BOOL | VT_BYREF) && cancel.pboolVal) *cancel.pboolVal = VARIANT_TRUE;
            }
            s->closePending = true;
        }
        return S_OK;
    }

private:
    LONG refs_ = 1;
    std::wstring key_;
    GUID iid_;
};

void resolveWorkbookDispIds(IDispatch* workbook, const GUID& iid) {
    ComPtr<ITypeInfo> info;
    if (FAILED(workbook->GetTypeInfo(0, LOCALE_USER_DEFAULT, &info)) || !info) return;
    ComPtr<ITypeLib> lib;
    UINT index = 0;
    if (FAILED(info->GetContainingTypeLib(&lib, &index))) return;
    ComPtr<ITypeInfo> events;
    if (FAILED(lib->GetTypeInfoOfGuid(iid, &events))) return;
    LPOLESTR names[] = {const_cast<LPOLESTR>(L"SheetChange"), const_cast<LPOLESTR>(L"BeforeClose")};
    DISPID id = 0;
    if (SUCCEEDED(events->GetIDsOfNames(&names[0], 1, &id))) g_sheetChange = id;
    if (SUCCEEDED(events->GetIDsOfNames(&names[1], 1, &id))) g_beforeClose = id;
}

// Abonniert SheetChange/BeforeClose der Arbeitsmappe; false, wenn Excel das nicht zulässt.
bool connectWorkbook(Sheet& s, const std::wstring& key) {
    ComPtr<IConnectionPointContainer> container;
    if (FAILED(s.workbook->QueryInterface(IID_PPV_ARGS(&container)))) return false;
    ComPtr<IEnumConnectionPoints> points;
    if (FAILED(container->EnumConnectionPoints(&points))) return false;
    ComPtr<IConnectionPoint> point;
    ULONG fetched = 0;
    if (points->Next(1, &point, &fetched) != S_OK) return false;
    GUID iid{};
    if (FAILED(point->GetConnectionInterface(&iid))) return false;
    resolveWorkbookDispIds(s.workbook.Get(), iid);
    auto* sink = new WorkbookSink(key, iid);
    const HRESULT hr = point->Advise(sink, &s.cookie);
    sink->Release();
    if (FAILED(hr)) {
        s.cookie = 0;
        return false;
    }
    s.point = point;
    return true;
}

ComPtr<IDispatch> excelApp() {
    if (g_excel) {
        try {
            disp::call(g_excel.Get(), L"Version");
        } catch (...) {
            g_excel.Reset();  // vom Nutzer beendet
        }
    }
    if (!g_excel) {
        CLSID clsid{};
        check(CLSIDFromProgID(L"Excel.Application", &clsid), "Excel ist nicht installiert");
        check(CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&g_excel)), "Excel starten");
    }
    return g_excel;
}

void bringToFront(Sheet& s) {
    auto excel = excelApp();
    disp::put(excel.Get(), L"Visible", true);
    disp::call(s.workbook.Get(), L"Activate");
    try {
        SetForegroundWindow(reinterpret_cast<HWND>(static_cast<LONG_PTR>(disp::call(excel.Get(), L"Hwnd").toLong())));
    } catch (...) {
    }
}

bool openInExcelFile(const fs::path& file) {
    const std::wstring quoted = L"\"" + file.wstring() + L"\"";
    if (reinterpret_cast<INT_PTR>(ShellExecuteW(nullptr, L"open", L"excel.exe", quoted.c_str(), nullptr, SW_SHOWNORMAL)) > 32) return true;
    return reinterpret_cast<INT_PTR>(ShellExecuteW(nullptr, L"open", file.c_str(), nullptr, nullptr, SW_SHOWNORMAL)) > 32;
}

void releaseSheet(Sheet& s) {
    if (s.point && s.cookie) s.point->Unadvise(s.cookie);
    s.point.Reset();
    s.cookie = 0;
    s.workbook.Reset();
}

void quitExcelIfUnused() {
    if (!g_excel) return;
    for (auto& [key, s] : g_sheets)
        if (s.workbook) return;
    try {
        if (disp::call(disp::object(g_excel.Get(), L"Workbooks").Get(), L"Count").toLong() == 0) disp::call(g_excel.Get(), L"Quit");
    } catch (...) {
    }
    g_excel.Reset();
}

// Prüft alle Datenblätter; liefert die Zahl neu gezeichneter Elemente.
int refreshAll(bool reportErrors) {
    if (g_busy) return 0;
    g_busy = true;
    int redrawn = 0;
    std::vector<std::wstring> finished;
    const ULONGLONG now = GetTickCount64();
    for (auto& [key, s] : g_sheets) {
        try {
            if (s.closed) {
                finished.push_back(key);
            } else if (s.closePending) {
                flushWorkbook(s);  // bei "beschäftigt" wird es im nächsten Takt erneut versucht
                s.dirty = false;
                ++redrawn;
                disp::put(s.workbook.Get(), L"Saved", true);
                s.closingByUs = true;
                disp::call(s.workbook.Get(), L"Close", false);
                s.closed = true;
                finished.push_back(key);
            } else if (s.cookie) {
                if (s.dirty && (now - s.dirtyAt >= 300 || !reportErrors)) {
                    flushWorkbook(s);
                    s.dirty = false;
                    ++redrawn;
                }
            } else {
                std::error_code ec;
                const auto stamp = fs::last_write_time(s.file, ec);
                if (ec) {
                    finished.push_back(key);
                } else if (stamp != s.stamp) {
                    const auto bytes = readFileShared(s.file);
                    if (!bytes.empty()) {
                        s.stamp = stamp;
                        applyTable(s, onWorker([&] { return tableFromDatasheet(bytes); }));
                        ++redrawn;
                    }
                }
            }
        } catch (const std::exception& e) {
            if (isBusy(e)) continue;  // Excel ist beschäftigt (z. B. Zelle im Bearbeitungsmodus): beim nächsten Takt erneut
            const std::wstring text = widen(e.what());
            if (text.find(L"nicht mehr auf der Folie") != std::wstring::npos) finished.push_back(key);
            s.dirty = false;
            std::error_code ec;
            s.stamp = fs::last_write_time(s.file, ec);  // dieselbe Fassung nicht erneut melden
            if (reportErrors) MessageBoxW(nullptr, (L"Datenblatt „" + s.name + L"“: " + text).c_str(), kTitle, MB_ICONWARNING);
        } catch (...) {
            finished.push_back(key);  // Folie/Präsentation nicht mehr erreichbar
        }
    }
    for (auto& key : finished) {
        if (auto* s = sheetFor(key)) {
            if (!s->closed && s->workbook) {
                try {
                    disp::put(s->workbook.Get(), L"Saved", true);
                    disp::call(s->workbook.Get(), L"Close", false);
                } catch (...) {
                }
            }
            releaseSheet(*s);
        }
        g_sheets.erase(key);
    }
    if (!finished.empty()) quitExcelIfUnused();
    if (g_sheets.empty() && g_timer) {
        KillTimer(nullptr, g_timer);
        g_timer = 0;
    }
    g_busy = false;
    return redrawn;
}

void CALLBACK onTimer(HWND, UINT, UINT_PTR, DWORD) { refreshAll(true); }

// Ein Datenblatt pro Element: Folien-ID + Elementname bleiben beim Neuzeichnen gleich (die Form-ID nicht).
fs::path datasheetPath(const std::wstring& name, long slideId) {
    wchar_t temp[MAX_PATH + 1] = {};
    GetTempPathW(MAX_PATH, temp);
    std::wstring safe;
    for (wchar_t c : name) safe += std::wstring(L"\\/:*?\"<>|").find(c) == std::wstring::npos ? c : L'_';
    const fs::path dir = fs::path(temp) / L"rethink";
    fs::create_directories(dir);
    return dir / std::format(L"Datenblatt {} (Folie {}).xlsx", safe, slideId);
}

std::wstring sheetKey(IDispatch* element) {
    auto slide = disp::object(element, L"Parent");
    return toLower(datasheetPath(disp::call(element, L"Name").str(), disp::call(slide.Get(), L"SlideID").toLong()).wstring());
}

}  // namespace

// ---- Öffentliche Funktionen ----

ComPtr<IDispatch> rethinkElementOf(IDispatch* selection) {
    if (!selection) return nullptr;
    try {
        const long type = disp::call(selection, L"Type").toLong();
        if (type != kPpSelectionShapes && type != kPpSelectionText) return nullptr;
        auto range = disp::object(selection, L"ShapeRange");
        if (disp::call(range.Get(), L"Count").toLong() != 1) return nullptr;
        auto shape = disp::object(range.Get(), L"Item", 1L);
        for (int depth = 0; depth < 2 && shape; ++depth) {
            if (disp::call(shape.Get(), L"Type").toLong() == kMsoGroup && !tagValue(shape.Get(), kModelTag).empty()) return shape;
            try {
                shape = disp::object(shape.Get(), L"ParentGroup");
            } catch (const Error&) {
                break;
            }
        }
    } catch (const Error&) {
    }
    return nullptr;
}

std::wstring insertElement(IDispatch* app, ElementKind kind, Orientation orientation, bool openSheet) {
    auto slide = currentSlide(app);
    auto pageSetup = disp::object(disp::object(activeWindow(app).Get(), L"Presentation").Get(), L"PageSetup");
    const double slideWidth = disp::call(pageSetup.Get(), L"SlideWidth").toDouble();
    const double slideHeight = disp::call(pageSetup.Get(), L"SlideHeight").toDouble();
    const bool gantt = kind == ElementKind::Gantt;
    const double width = slideWidth * (gantt ? 0.84 : 0.7), height = slideHeight * 0.6;

    const std::wstring name = freeName(slide.Get(), gantt ? L"Gantt" : L"Wasserfall");
    const Box box{emu((slideWidth - width) / 2), emu((slideHeight - height) / 2 + slideHeight * 0.05), emu(width), emu(height)};
    const DataItem item = gantt ? sampleGantt(name) : sampleWaterfall(name, kind == ElementKind::WaterfallDown, orientation);
    auto group = renderElement(slide.Get(), nullptr, item, box);
    try {
        disp::call(group.Get(), L"Select");
    } catch (const Error&) {
    }
    if (openSheet) openDatasheet(app, true);
    return name;
}

std::wstring insertWaterfall(IDispatch* app, bool openSheet) { return insertElement(app, ElementKind::Waterfall, Orientation::Up, openSheet); }

ElementInfo elementInfo(IDispatch* element) {
    ElementInfo info;
    const std::wstring kind = tagValue(element, kKindTag);  // ältere Elemente ohne Tag: Wasserfall nach oben
    info.gantt = kind == L"gantt";
    for (int o = 0; o < 4; ++o)
        if (kind == std::wstring(L"waterfall ") + orientationName(static_cast<Orientation>(o))) info.direction = static_cast<Orientation>(o);
    return info;
}

// ---- Eingebautes Datenblatt ----

namespace {

void sortCategories(IDispatch* app, IDispatch* element, const std::wstring& order);  // Abschnitt Zeitplan

struct GridContext {
    ComPtr<IDispatch> slide;
    long shapeId = 0;
    std::wstring name;
    Datasheet* window = nullptr;
};
std::map<std::wstring, GridContext*> g_grids;  // Schlüssel wie beim Excel-Datenblatt (Folie + Name)
ComPtr<IDispatch> g_gridApp;

HWND powerPointFrame() {
    HWND found = nullptr;
    EnumThreadWindows(
        GetCurrentThreadId(),
        [](HWND hwnd, LPARAM param) -> BOOL {
            wchar_t name[64] = {};
            GetClassNameW(hwnd, name, 64);
            if (wcscmp(name, L"PPTFrameClass") == 0 && IsWindowVisible(hwnd)) {
                *reinterpret_cast<HWND*>(param) = hwnd;
                return FALSE;
            }
            return TRUE;
        },
        reinterpret_cast<LPARAM>(&found));
    return found ? found : GetActiveWindow();
}

RECT screenRectOf(IDispatch* app, IDispatch* element) {
    auto window = disp::object(app, L"ActiveWindow");
    const Rect r = rectOf(element);
    return {disp::call(window.Get(), L"PointsToScreenPixelsX", r.left).toLong(), disp::call(window.Get(), L"PointsToScreenPixelsY", r.top).toLong(),
            disp::call(window.Get(), L"PointsToScreenPixelsX", r.left + r.width).toLong(),
            disp::call(window.Get(), L"PointsToScreenPixelsY", r.top + r.height).toLong()};
}

// Eingabe im Datenblatt → Element an Ort und Stelle neu zeichnen.
void applyGridTable(GridContext& g, const std::vector<Row>& table) {
    try {
        auto group = findShape(g.slide.Get(), g.shapeId);
        if (!group) throw Error(std::format("Das Diagramm '{}' ist nicht mehr auf der Folie.", narrow(g.name)));
        const std::wstring stored = tagValue(group.Get(), kModelTag);
        DataItem item = onWorker([&] {
            DataItem parsed = parseDataItemJson(stored);
            parsed.table = table;
            return parsed;
        });
        auto result = renderElement(g.slide.Get(), group, item, layoutBox(group.Get()));
        g.shapeId = disp::call(result.Get(), L"Id").toLong();
    } catch (const std::exception& e) {
        MessageBoxW(g.window ? datasheetWindow(g.window) : nullptr, widen(e.what()).c_str(), kTitle, MB_ICONWARNING);
    }
}

GridContext* gridFor(IDispatch* element) {
    auto it = g_grids.find(sheetKey(element));
    return it == g_grids.end() ? nullptr : it->second;
}

// Datenblatt des ausgewählten Elements; ohne passende Auswahl das einzige offene Datenblatt.
GridContext& requireGrid(IDispatch* app) {
    ComPtr<IDispatch> element;
    try {
        element = rethinkElementOf(disp::object(activeWindow(app).Get(), L"Selection").Get());
    } catch (const Error&) {
    }
    if (element)
        if (GridContext* g = gridFor(element.Get())) return *g;
    if (g_grids.size() == 1) return *g_grids.begin()->second;
    throw Error("Für das ausgewählte Diagramm ist kein Datenblatt geöffnet.");
}

}  // namespace

void openDatasheet(IDispatch* app, bool show) {
    auto element = selectedElement(app);
    const std::wstring key = sheetKey(element.Get());
    if (auto it = g_grids.find(key); it != g_grids.end()) {
        it->second->shapeId = disp::call(element.Get(), L"Id").toLong();
        if (show) datasheetActivate(it->second->window);
        return;
    }
    const std::wstring stored = tagValue(element.Get(), kModelTag);
    std::vector<Row> table = onWorker([&] { return parseDataItemJson(stored).table; });

    auto* ctx = new GridContext{disp::object(element.Get(), L"Parent"), disp::call(element.Get(), L"Id").toLong(),
                                disp::call(element.Get(), L"Name").str(), nullptr};
    g_gridApp = app;
    DatasheetCallbacks callbacks;
    callbacks.changed = [ctx](const std::vector<Row>& t) { applyGridTable(*ctx, t); };
    callbacks.openInExcel = [ctx]() {
        try {
            if (g_gridApp) {
                if (auto group = findShape(ctx->slide.Get(), ctx->shapeId)) disp::call(group.Get(), L"Select");
                openExcelDatasheet(g_gridApp.Get(), true);
            }
        } catch (const std::exception& e) {
            MessageBoxW(nullptr, widen(e.what()).c_str(), kTitle, MB_ICONWARNING);
        }
    };
    if (elementInfo(element.Get()).gantt)
        callbacks.sortCategories = [ctx]() {
            try {
                auto group = g_gridApp ? findShape(ctx->slide.Get(), ctx->shapeId) : nullptr;
                if (!group) throw Error(std::format("Das Diagramm '{}' ist nicht mehr auf der Folie.", narrow(ctx->name)));
                sortCategories(g_gridApp.Get(), group.Get(), L"");  // führt dieses Datenblatt mit
                ctx->shapeId = disp::call(disp::object(disp::object(ctx->slide.Get(), L"Shapes").Get(), L"Item", ctx->name).Get(), L"Id").toLong();
            } catch (const std::exception& e) {
                MessageBoxW(ctx->window ? datasheetWindow(ctx->window) : nullptr, widen(e.what()).c_str(), kTitle, MB_ICONWARNING);
            }
        };
    callbacks.closed = [key, ctx]() {
        g_grids.erase(key);
        delete ctx;
    };
    const std::wstring hint = elementInfo(element.Get()).gantt
                                  ? L"Zeile 1: Überschriften · Zeile 2: Meilensteinzeile · darunter je Aktivität: Start, Ende (00:00 Uhr), Meilenstein, "
                                    L"Verantwortlich, Bemerkung · Datum als TT.MM.JJJJ"
                                  : L"";
    ctx->window = openDatasheetWindow(powerPointFrame(), L"rethink – Datenblatt " + ctx->name, std::move(table), localeDecimalSeparator(),
                                      std::move(callbacks), screenRectOf(app, element.Get()), show, hint);
    g_grids[key] = ctx;
}

void datasheetEnterCell(IDispatch* app, int row, int column, const std::wstring& text) {
    datasheetEnter(requireGrid(app).window, row, column, text);
}

void datasheetPasteText(IDispatch* app, int row, int column, const std::wstring& text) {
    datasheetPaste(requireGrid(app).window, row, column, text);
}

std::wstring datasheetCellText(IDispatch* app, int row, int column) { return datasheetText(requireGrid(app).window, row, column); }

void closeDatasheet(IDispatch* app) { datasheetClose(requireGrid(app).window); }

int openDatasheetCount() { return static_cast<int>(g_grids.size()); }

std::wstring openExcelDatasheet(IDispatch* app, bool show) {
    auto element = selectedElement(app);
    auto slide = disp::object(element.Get(), L"Parent");
    const long shapeId = disp::call(element.Get(), L"Id").toLong();
    const std::wstring name = disp::call(element.Get(), L"Name").str();
    const fs::path file = datasheetPath(name, disp::call(slide.Get(), L"SlideID").toLong());
    const std::wstring key = toLower(file.wstring());

    if (Sheet* existing = sheetFor(key); existing && !existing->closed) {
        existing->shapeId = shapeId;  // bereits geöffnet: dieselbe Mappe nach vorn holen
        existing->slide = slide;
        if (show) {
            if (existing->workbook)
                bringToFront(*existing);
            else
                openInExcelFile(file);
        }
        return file.wstring();
    }

    const std::wstring stored = tagValue(element.Get(), kModelTag);
    const auto bytes = onWorker([&] { return datasheetFromTable(parseDataItemJson(stored).table); });
    HANDLE h = CreateFileW(file.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE)
        throw Error("Das Datenblatt ist noch in Excel geöffnet. Bitte dort weiterarbeiten oder die Datei schließen: " + narrow(file.wstring()));
    DWORD written = 0;
    WriteFile(h, bytes.data(), static_cast<DWORD>(bytes.size()), &written, nullptr);
    CloseHandle(h);

    Sheet s;
    s.file = file;
    s.stamp = fs::last_write_time(file);
    s.slide = slide;
    s.shapeId = shapeId;
    s.name = name;
    g_sheets[key] = s;
    Sheet& sheet = g_sheets[key];

    // Bevorzugt: eigene Excel-Instanz per Automation (Änderungen live, Schließen ohne Rückfrage).
    try {
        auto excel = excelApp();
        sheet.workbook = disp::object(disp::object(excel.Get(), L"Workbooks").Get(), L"Open", file.wstring());
        try {
            disp::put(disp::object(disp::object(sheet.workbook.Get(), L"Windows").Get(), L"Item", 1L).Get(), L"Caption",
                      L"rethink – Datenblatt " + name);
        } catch (...) {
        }
        if (!connectWorkbook(sheet, key)) sheet.workbook.Reset();
    } catch (...) {
        sheet.workbook.Reset();
    }
    if (!g_timer) g_timer = SetTimer(nullptr, 0, 300, onTimer);

    if (sheet.workbook) {
        if (show) bringToFront(sheet);
    } else if (show && !openInExcelFile(file)) {
        throw Error("Excel konnte das Datenblatt nicht öffnen: " + narrow(file.wstring()));
    }
    return file.wstring();
}

ComPtr<IDispatch> datasheetWorkbook(IDispatch* app) {
    Sheet* s = sheetFor(sheetKey(selectedElement(app).Get()));
    return s ? s->workbook : nullptr;
}

int applyDatasheets(IDispatch*) { return refreshAll(false); }

namespace {

// Zelle hinter einem Segment; wirft, wenn sie sich nicht direkt ändern lässt.
Cell& valueCell(DataItem& item, int row, int column) {
    if (row < 1 || static_cast<size_t>(row) >= item.table.size() || column < 1 || static_cast<size_t>(column) >= item.table[row].size())
        throw Error("Dieses Segment gibt es nicht mehr.");
    Cell& cell = item.table[row][column];
    if (isComputedCell(cell)) throw Error("Dieses Segment wird aus „e“ berechnet und kann nicht direkt verändert werden.");
    if (cell.kind != CellKind::Number && cell.kind != CellKind::Percentage) throw Error("Dieses Segment hat keinen Zahlenwert.");
    return cell;
}

// Wert in die Zelle des offenen Excel-Datenblatts schreiben (Excel speichert Prozente als Anteil).
void writeExcelCell(Sheet& s, int row, int column, const Cell& cell) {
    try {
        auto sheet = disp::object(disp::object(s.workbook.Get(), L"Worksheets").Get(), L"Item", 1L);
        auto range = disp::object(sheet.Get(), L"Range", columnName(static_cast<size_t>(column)) + std::to_wstring(row + 1));
        disp::put(range.Get(), L"Value2", cell.kind == CellKind::Percentage ? cell.number / 100.0 : cell.number);
    } catch (const std::exception& e) {
        if (isBusy(e)) throw Error("Excel ist gerade beschäftigt (Zelle im Bearbeitungsmodus?). Bitte die Zelle verlassen und erneut klicken.");
        throw;
    }
}

Sheet* openExcelSheet(IDispatch* element) {
    Sheet* s = sheetFor(sheetKey(element));
    return s && s->workbook && !s->closed ? s : nullptr;
}

}  // namespace

void adjustSegment(IDispatch* app, int row, int column, int direction, bool large) {
    auto element = selectedElement(app);
    auto slide = disp::object(element.Get(), L"Parent");
    const std::wstring stored = tagValue(element.Get(), kModelTag);

    struct Change {
        DataItem item;
        Cell cell;
    };
    const Change change = onWorker([&] {
        Change c;
        c.item = parseDataItemJson(stored);
        Cell& cell = valueCell(c.item, row, column);
        // Schrittweite nach der Genauigkeit des Werts selbst (12 → 1, 12,5 → 0,1), nicht nach dem ganzen Diagramm
        const int decimals = stepDecimals(cell, parseChartOptions(c.item.chartJson));
        const double scale = std::pow(10.0, decimals);
        const double step = (large ? 10.0 : 1.0) / scale;
        cell.number = std::round((cell.number + direction * step) * scale) / scale;
        c.cell = cell;
        return c;
    });

    // Ist das Datenblatt offen, ist es die Quelle: die Zelle ändern, das Änderungsereignis zeichnet neu.
    if (Sheet* s = openExcelSheet(element.Get())) {
        writeExcelCell(*s, row, column, change.cell);
        s->dirty = true;
        s->dirtyAt = GetTickCount64() - 300;  // sofort übernehmen
        refreshAll(false);
        return;
    }
    renderElement(slide.Get(), element, change.item, layoutBox(element.Get()));
    if (GridContext* g = gridFor(element.Get())) datasheetSetTable(g->window, change.item.table);  // offenes Datenblatt mitführen
}

// ---- Ziehen der ▲▼-Griffe ----

namespace {

// Zustand eines laufenden Ziehens. Modell, Optionen und Layout-Rahmen werden zu Beginn eingefroren; jeder Schritt
// rechnet nur das Layout neu (reines C++, ohne Worker-Thread) und fasst nur Formen an, die sich geändert haben.
// Der Vergleich mit der letzten Zeichnung ist nur hier zulässig: während des Ziehens hält das Overlay die Maus,
// niemand sonst kann die Formen verändern.
struct SegmentDrag {
    ComPtr<IDispatch> slide, group;
    DataItem item;
    ChartOptions options;
    Box box;
    int row = 0, column = 0;
    Drawing last;         // zuletzt gezeichnet
    bool synced = false;  // Formen entsprechen `last` (erst nach dem ersten Schritt)
    RoleMap byRole;
};
std::optional<SegmentDrag> g_drag;

Drawing layoutForDrag(const SegmentDrag& drag) {
    try {
        return layoutChart(drag.item, drag.options, drag.box);
    } catch (const std::invalid_argument& e) {
        throw Error(std::format("Diagramm '{}': {}", narrow(drag.item.name), e.what()));
    }
}

std::vector<SegmentInfo> segmentsOfDrawing(const Drawing& d) {
    std::vector<SegmentInfo> result;
    for (auto& p : d.prims) {
        SegmentInfo s;
        if (p.kind != Prim::Kind::Rect || sscanf_s(p.role.c_str(), "segment r%d c%d", &s.row, &s.column) != 2) continue;
        s.flags = p.flags;
        s.left = pt(p.x), s.top = pt(p.y), s.width = pt(p.w), s.height = pt(p.h);
        result.push_back(s);
    }
    return result;
}

bool sameRoles(const Drawing& a, const Drawing& b) {
    if (a.prims.size() != b.prims.size()) return false;
    for (size_t i = 0; i < a.prims.size(); ++i)
        if (a.prims[i].role != b.prims[i].role) return false;
    return true;
}

}  // namespace

SegmentDragStart beginSegmentDrag(IDispatch* app, int row, int column) {
    g_drag.reset();
    auto element = selectedElement(app);
    const std::wstring stored = tagValue(element.Get(), kModelTag);

    SegmentDrag drag;
    drag.slide = disp::object(element.Get(), L"Parent");
    drag.group = element;
    drag.row = row, drag.column = column;
    drag.box = layoutBox(element.Get());
    std::tie(drag.item, drag.options) = onWorker([&] {
        DataItem item = parseDataItemJson(stored);
        ChartOptions options = optionsOf(item);
        return std::pair{std::move(item), std::move(options)};
    });
    const Cell& cell = valueCell(drag.item, row, column);

    SegmentDragStart start;
    start.value = cell.number;
    start.decimals = stepDecimals(cell, drag.options);
    drag.last = layoutForDrag(drag);
    start.pointsPerUnit = emuPerUnit(drag.last) / kEmuPerPt;

    auto byRole = shapesByRole(element.Get(), drag.last);
    if (!byRole) {  // Formen passen nicht mehr zum Modell (z. B. von Hand gelöscht): einmal vollständig neu zeichnen
        drag.group = renderElement(drag.slide.Get(), element, drag.item, drag.box);
        byRole = shapesByRole(drag.group.Get(), drag.last);
        if (!byRole) throw Error("Das Diagramm lässt sich nicht ziehen. Bitte das Datenblatt verwenden.");
        drag.synced = true;
    }
    drag.byRole = std::move(*byRole);

    // Alle Änderungen des Ziehens bilden einen einzigen Rückgängig-Eintrag (ab PowerPoint 2010).
    try {
        disp::call(app, L"StartNewUndoEntry");
    } catch (const std::exception&) {
    }
    g_drag = std::move(drag);
    return start;
}

bool dragSegmentTo(double value, std::vector<SegmentInfo>& segments) {
    if (!g_drag) return false;
    SegmentDrag& drag = *g_drag;
    Cell& cell = drag.item.table[drag.row][drag.column];
    if (drag.synced && cell.number == value) return false;
    cell.number = value;

    Drawing d = layoutForDrag(drag);
    if (!sameRoles(d, drag.last)) {
        // Andere Formen (selten, z. B. Beschriftung entfällt): vollständig neu zeichnen und Rollen neu einlesen
        drag.group = renderElement(drag.slide.Get(), drag.group, drag.item, drag.box);
        auto byRole = shapesByRole(drag.group.Get(), d);
        if (!byRole) throw Error("Das Diagramm lässt sich nicht weiter ziehen.");
        drag.byRole = std::move(*byRole);
    } else {
        for (size_t i = 0; i < d.prims.size(); ++i)
            if (!drag.synced || !samePrim(d.prims[i], drag.last.prims[i])) updateShape(drag.byRole[widen(d.prims[i].role)].Get(), d.prims[i]);
    }
    drag.last = std::move(d);
    drag.synced = true;
    segments = segmentsOfDrawing(drag.last);
    return true;
}

void endSegmentDrag() {
    if (!g_drag) return;
    SegmentDrag drag = std::move(*g_drag);
    g_drag.reset();
    if (!drag.synced) return;  // Wert nie geändert

    const DataItem& item = drag.item;
    setTag(drag.group.Get(), kModelTag, onWorker([&] { return serializeDataItem(item); }));
    storeLayout(drag.group.Get(), drag.box);
    if (GridContext* g = gridFor(drag.group.Get())) datasheetSetTable(g->window, item.table);
    // Excel-Datenblatt: Zelle nachführen; das Änderungsereignis zeichnet mit denselben Werten (keine Änderung).
    if (Sheet* s = openExcelSheet(drag.group.Get())) writeExcelCell(*s, drag.row, drag.column, item.table[drag.row][drag.column]);
}

// ---- Zeitplan (Gantt) ----

namespace {

using GanttChange = std::function<void(DataItem&, ChartOptions&)>;

ComPtr<IDispatch> selectedGantt(IDispatch* app) {
    auto element = selectedElement(app);
    if (!elementInfo(element.Get()).gantt) throw Error("Bitte einen rethink-Zeitplan auswählen.");
    return element;
}

// Ganze Tabelle in das offene Excel-Datenblatt schreiben (Datumsangaben als Excel-Datum).
void writeExcelTable(Sheet& s, const std::vector<Row>& table) {
    auto sheet = disp::object(disp::object(s.workbook.Get(), L"Worksheets").Get(), L"Item", 1L);
    disp::call(disp::object(sheet.Get(), L"UsedRange").Get(), L"ClearContents");
    for (size_t r = 0; r < table.size(); ++r)
        for (size_t c = 0; c < table[r].size(); ++c) {
            const Cell& cell = table[r][c];
            if (cell.kind == CellKind::Empty) continue;
            auto range = disp::object(sheet.Get(), L"Range", columnName(c) + std::to_wstring(r + 1));
            switch (cell.kind) {
                case CellKind::Number: disp::put(range.Get(), L"Value2", cell.number); break;
                case CellKind::Percentage: disp::put(range.Get(), L"Value2", cell.number / 100.0); break;
                case CellKind::String: disp::put(range.Get(), L"Value2", cell.text); break;
                case CellKind::Date:
                    if (const auto t = cellDate(cell)) {
                        disp::put(range.Get(), L"NumberFormat", std::wstring(L"dd.mm.yyyy"));
                        disp::put(range.Get(), L"Value2", *t + 25569.0);  // Excel zählt ab 30.12.1899
                    }
                    break;
                case CellKind::Empty: break;
            }
        }
}

void syncDatasheets(IDispatch* group, const std::vector<Row>& table) {
    if (GridContext* g = gridFor(group)) datasheetSetTable(g->window, table);
    if (Sheet* s = openExcelSheet(group)) {
        try {
            writeExcelTable(*s, table);
        } catch (const std::exception&) {
            // Excel beschäftigt: das Diagramm stimmt, das Datenblatt holt beim nächsten Öffnen auf
        }
    }
}

// Änderungen über Zeilen, Elemente und Optionen: Modell lesen, ändern, zurückschreiben, an Ort und Stelle zeichnen.
ComPtr<IDispatch> modifyGantt(IDispatch* app, IDispatch* element, const GanttChange& change) {
    auto slide = disp::object(element, L"Parent");
    const std::wstring stored = tagValue(element, kModelTag);
    const DataItem item = onWorker([&] {
        DataItem parsed = parseDataItemJson(stored);
        ChartOptions options = optionsOf(parsed);
        try {
            change(parsed, options);
        } catch (const std::invalid_argument& e) {
            throw Error(e.what());
        }
        parsed.chartJson = mergeChartJson(parsed.chartJson, ganttEditsJson(options.gantt));
        return parsed;
    });
    try {
        disp::call(app, L"StartNewUndoEntry");
    } catch (const std::exception&) {
    }
    auto group = renderElement(slide.Get(), element, item, layoutBox(element));
    syncDatasheets(group.Get(), item.table);
    return group;
}

struct GanttModel {
    DataItem item;
    ChartOptions options;
    Drawing drawing;
};

GanttModel ganttModel(IDispatch* element) {
    const std::wstring stored = tagValue(element, kModelTag);
    GanttModel m;
    std::tie(m.item, m.options) = onWorker([&] {
        DataItem item = parseDataItemJson(stored);
        ChartOptions options = optionsOf(item);
        return std::pair{std::move(item), std::move(options)};
    });
    m.drawing = layoutOrThrow(m.item, m.options, layoutBox(element));
    return m;
}

bool draggable(const std::string& flags) { return flags == "bar" || flags == "milestone" || flags == "bracket" || flags == "shade"; }

GanttGeometry geometryOfDrawing(const Drawing& d) {
    GanttGeometry g;
    g.valid = d.axis.valid;
    for (auto& p : d.prims) {
        if (!draggable(p.flags)) continue;
        GanttShape s;
        s.role = p.role, s.flags = p.flags;
        s.left = pt(std::min(p.x, p.x + p.w)), s.top = pt(std::min(p.y, p.y + p.h)), s.width = pt(std::abs(p.w)), s.height = pt(std::abs(p.h));
        g.shapes.push_back(s);
    }
    g.plotLeft = pt(d.axis.x0), g.plotRight = pt(d.axis.x0 + d.axis.w), g.scaleTop = pt(d.scaleTop), g.scaleBottom = pt(d.scaleBottom);
    for (auto& r : d.rows)
        if (r.row >= gantt::kFirstActivityRow) {
            if (g.rowCount++ == 0) g.rowsTop = pt(r.y);
            g.rowH = pt(r.h);
        }
    g.labelsLeft = pt(d.bounds.x);
    return g;
}

std::wstring datesText(const GanttTarget& t) {
    if (t.kind == GanttTarget::Kind::Milestone) return formatDate(t.start, L"ddd dd.MM.yyyy", L"de");
    return formatDate(t.start, L"dd.MM.", L"de") + L"–" + formatDate(std::ceil(t.end - 1e-9) - 1, L"dd.MM.yyyy", L"de");
}

// Zustand eines Ziehens am Zeitplan; wie beim Wasserfall wird nur das Layout neu gerechnet und nur Geändertes angefasst.
struct GanttDrag {
    ComPtr<IDispatch> app, slide, group;
    DataItem item0, item;
    ChartOptions options0, options;
    Box box;
    Drawing last;
    RoleMap byRole;
    bool range = false;
    GanttTarget target;
    GanttGrip grip = GanttGrip::Move;
    TimeUnit unit = TimeUnit::Day;
    double first = 0, second = 0;  // aktueller Stand (Bereich: t0/t1, sonst Start/Ende)
    bool changed = false;
    std::wstring autoStart, autoEnd;  // gespeicherter Bereich (leer = folgt den Daten), nach dem Ziehen wiederhergestellt
};
std::optional<GanttDrag> g_ganttDrag;

std::wstring dragTooltip(const GanttDrag& g) {
    const std::wstring& lang = g.options.language;
    if (g.range) return ganttTooltip(g.grip == GanttGrip::Start ? g.first : g.second - 1, lang);
    if (g.target.kind == GanttTarget::Kind::Milestone) return ganttTooltip(g.first, lang);
    switch (g.grip) {
        case GanttGrip::Start: return ganttTooltip(g.first, lang);
        case GanttGrip::End: return ganttTooltip(std::ceil(g.second - 1e-9) - 1, lang);  // letzter eingeschlossener Tag
        case GanttGrip::Move: {
            GanttTarget t = g.target;
            t.start = g.first, t.end = g.second;
            return datesText(t);
        }
    }
    return {};
}

}  // namespace

GanttGeometry ganttGeometryOf(IDispatch* element) {
    GanttGeometry g;
    auto items = disp::object(element, L"GroupItems");
    const long count = disp::call(items.Get(), L"Count").toLong();
    for (long i = 1; i <= count; ++i) {
        auto shape = disp::object(items.Get(), L"Item", i);
        const std::string flags = narrow(tagValue(shape.Get(), kSegmentTag));
        if (!draggable(flags)) continue;
        const Rect r = rectOf(shape.Get());
        g.shapes.push_back({narrow(disp::call(shape.Get(), L"Name").str()), flags, r.left, r.top, r.width, r.height});
    }
    // Zeitachse und Skalenband aus dem Tag beim letzten Zeichnen, um die seither verschobene Gruppe versetzt
    std::wistringstream in(tagValue(element, kGanttTag));
    double plotLeft = 0, plotRight = 0, scaleTop = 0, scaleBottom = 0, rowsTop = 0, rowH = 0, left = 0;
    int rows = 0;
    wchar_t sep = 0;
    in >> plotLeft >> sep >> plotRight >> sep >> scaleTop >> sep >> scaleBottom >> sep >> rowsTop >> sep >> rowH >> sep >> rows >> sep >> left;
    const auto layout = storedLayout(element);
    if (!in.fail() && layout) {
        const Rect now = rectOf(element);
        const double dx = now.left - layout->group.left, dy = now.top - layout->group.top;
        g.plotLeft = plotLeft + dx, g.plotRight = plotRight + dx, g.scaleTop = scaleTop + dy, g.scaleBottom = scaleBottom + dy;
        g.rowsTop = rowsTop + dy, g.rowH = rowH, g.rowCount = rows, g.labelsLeft = left + dx;
        g.valid = true;
    }
    return g;
}

std::wstring beginGanttDrag(IDispatch* app, const std::string& role, int grip, double pointsPerPixel) {
    g_ganttDrag.reset();
    auto element = selectedGantt(app);
    GanttModel m = ganttModel(element.Get());
    GanttDrag drag;
    drag.app = app;
    drag.slide = disp::object(element.Get(), L"Parent");
    drag.group = element;
    drag.box = layoutBox(element.Get());
    drag.grip = static_cast<GanttGrip>(std::clamp(grip, 0, 2));
    drag.unit = m.drawing.axis.snapUnit(3 * std::max(pointsPerPixel, 0.01) * kEmuPerPt);
    if (role == "range") {
        drag.range = true;
        if (drag.grip == GanttGrip::Move) drag.grip = GanttGrip::End;
        drag.first = m.drawing.axis.t0, drag.second = m.drawing.axis.t1;
        drag.unit = m.drawing.scales.empty() ? TimeUnit::Day : m.drawing.scales.back();  // Schritte der feinsten Skala
    } else {
        drag.target = ganttTarget(m.item, m.options, role);
        if (drag.target.kind == GanttTarget::Kind::None) throw Error("Dieses Element lässt sich nicht ziehen.");
        if (drag.target.kind == GanttTarget::Kind::Milestone) drag.grip = GanttGrip::Move;
        drag.first = drag.target.start, drag.second = drag.target.end;
        auto byRole = shapesByRole(element.Get(), m.drawing);
        if (!byRole) {  // Formen passen nicht mehr zum Modell: einmal vollständig neu zeichnen
            drag.group = renderElement(drag.slide.Get(), element, m.item, drag.box);
            byRole = shapesByRole(drag.group.Get(), m.drawing);
            if (!byRole) throw Error("Der Zeitplan lässt sich nicht ziehen. Bitte das Datenblatt verwenden.");
        }
        drag.byRole = std::move(*byRole);
    }
    drag.item0 = drag.item = m.item;
    drag.options0 = drag.options = m.options;
    // Während des Ziehens bleibt der sichtbare Bereich fest, auch wenn er sonst den Daten folgt
    drag.autoStart = m.options.gantt.start, drag.autoEnd = m.options.gantt.end;
    if (!drag.range) setGanttRange(drag.options0, m.drawing.axis.t0, m.drawing.axis.t1);
    drag.last = std::move(m.drawing);
    try {
        disp::call(app, L"StartNewUndoEntry");
    } catch (const std::exception&) {
    }
    const std::wstring tooltip = dragTooltip(drag);
    g_ganttDrag = std::move(drag);
    return tooltip;
}

bool ganttDragTo(double dx, GanttGeometry& geometry, std::wstring& tooltip) {
    if (!g_ganttDrag) return false;
    GanttDrag& g = *g_ganttDrag;
    const double dxEmu = dx * kEmuPerPt;
    if (g.range) {
        const auto [t0, t1] = ganttRangeDrag(g.last.axis, g.grip == GanttGrip::Start, dxEmu, g.unit);
        if (t0 == g.first && t1 == g.second) return false;
        g.first = t0, g.second = t1, g.changed = true;
        tooltip = dragTooltip(g);
        return true;  // gezeichnet wird erst beim Loslassen (die Skalen ändern sich vollständig)
    }
    const auto [s, e] = ganttDragDates(g.last.axis, g.target, g.grip, dxEmu, g.unit);
    if (s == g.first && e == g.second) return false;
    g.first = s, g.second = e, g.changed = true;
    g.item = g.item0;
    g.options = g.options0;
    engine::setGanttDates(g.item, g.options, g.target, s, e);
    Drawing d = layoutOrThrow(g.item, g.options, g.box);
    // Die Zeitachse bleibt beim Ziehen fest: neue Rollen (z. B. Pfeil am Bereichsende) erzwingen ein volles Zeichnen
    if (!sameRoles(d, g.last)) {
        g.group = renderElement(g.slide.Get(), g.group, g.item, g.box);
        auto byRole = shapesByRole(g.group.Get(), d);
        if (!byRole) throw Error("Der Zeitplan lässt sich nicht weiter ziehen.");
        g.byRole = std::move(*byRole);
    } else {
        for (size_t i = 0; i < d.prims.size(); ++i)
            if (!samePrim(d.prims[i], g.last.prims[i])) updateShape(g.byRole[widen(d.prims[i].role)].Get(), d.prims[i]);
    }
    const GanttAxis axis = g.last.axis;  // Maßstab des Ziehens bleibt der vom Anfang
    g.last = std::move(d);
    g.last.axis = axis;
    geometry = geometryOfDrawing(g.last);
    tooltip = dragTooltip(g);
    return true;
}

void endGanttDrag(bool cancel) {
    if (!g_ganttDrag) return;
    GanttDrag g = std::move(*g_ganttDrag);
    g_ganttDrag.reset();
    if (!g.changed) return;
    if (g.range) {
        if (cancel) return;
        const double t0 = g.first, t1 = g.second;
        modifyGantt(g.app.Get(), g.group.Get(), [&](DataItem&, ChartOptions& o) { setGanttRange(o, t0, t1); });
        return;
    }
    if (cancel) {  // Ausgangsstand wiederherstellen
        renderElement(g.slide.Get(), g.group, g.item0, g.box);
        return;
    }
    ChartOptions options = g.options;
    options.gantt.start = g.autoStart, options.gantt.end = g.autoEnd;
    DataItem item = g.item;
    item.chartJson = onWorker([&] { return mergeChartJson(item.chartJson, ganttEditsJson(options.gantt)); });
    if (g.autoStart.empty() || g.autoEnd.empty()) {
        renderElement(g.slide.Get(), g.group, item, g.box);  // Bereich folgt den Daten: einmal mit neuem Bereich zeichnen
    } else {
        setTag(g.group.Get(), kModelTag, onWorker([&] { return serializeDataItem(item); }));
        storeLayout(g.group.Get(), g.box);
    }
    syncDatasheets(g.group.Get(), item.table);
}

GanttHit ganttHitAt(IDispatch* element, double x, double y) {
    GanttHit hit;
    if (!elementInfo(element).gantt) return hit;
    const GanttModel m = ganttModel(element);
    const Drawing& d = m.drawing;
    const Emu ex = emu(x), ey = emu(y);
    hit.valid = true;
    for (auto& band : d.rows)
        if (ey >= band.y && ey < band.y + band.h) hit.row = band.row;
    if (ex >= d.axis.x0 && ex <= d.axis.x0 + d.axis.w) {
        hit.date = d.axis.dateAt(static_cast<double>(ex));
        hit.hasDate = true;
    }
    hit.inScale = hit.hasDate && ey >= d.scaleTop && ey < d.scaleBottom;
    const Emu tolerance = emu(3);
    for (auto it = d.prims.rbegin(); it != d.prims.rend(); ++it) {  // oberste Form zuerst
        if (!draggable(it->flags)) continue;
        const Emu x0 = std::min(it->x, it->x + it->w) - tolerance, x1 = std::max(it->x, it->x + it->w) + tolerance;
        const Emu y0 = std::min(it->y, it->y + it->h) - tolerance, y1 = std::max(it->y, it->y + it->h) + tolerance;
        if (ex >= x0 && ex <= x1 && ey >= y0 && ey <= y1) {
            hit.role = it->role;
            break;
        }
    }
    return hit;
}

namespace {

int requireRow(const GanttHit& hit) {
    if (hit.row < gantt::kFirstActivityRow) throw Error("Bitte mit der rechten Maustaste auf eine Zeile des Zeitplans klicken.");
    return hit.row;
}

double requireDate(const GanttHit& hit) {
    if (!hit.hasDate) throw Error("Bitte mit der rechten Maustaste im Zeitbereich des Zeitplans klicken.");
    return hit.date;
}

// Dauer neuer Balken: eine Einheit der feinsten Skala, mindestens eine Woche
double newBarDays(IDispatch* element) {
    const Drawing d = ganttModel(element).drawing;
    const TimeUnit unit = d.scales.empty() ? TimeUnit::Week : d.scales.back();
    return std::max(7.0, std::round(unitDays(unit)));
}

void addItem(IDispatch* app, IDispatch* element, GanttItem::Kind kind, const GanttHit& hit) {
    const double date = requireDate(hit);
    const double days = newBarDays(element);
    int row = hit.row;
    if (kind == GanttItem::Kind::Bar || kind == GanttItem::Kind::Pentagon || kind == GanttItem::Kind::Chevron) row = requireRow(hit);
    if (row < 0) row = gantt::kMilestoneRow;  // Meilensteine und Klammern unter dem Diagramm
    modifyGantt(app, element, [&](DataItem& item, ChartOptions& o) {
        addGanttItem(item, o, kind, row, date, kind == GanttItem::Kind::Shade ? 7 : days);
        if (kind == GanttItem::Kind::Pentagon || kind == GanttItem::Kind::Chevron) o.gantt.items.back().text = L"Phase";
    });
}

}  // namespace

namespace {

// IDs des PowerPoint-Kontextmenüs → Befehle
std::wstring commandOfId(const std::wstring& id) {
    static const std::map<std::wstring, std::wstring> kIds = {
        {L"rtkG_insertRow", L"insertRow"},     {L"rtkG_duplicateRow", L"duplicateRow"}, {L"rtkG_deleteRow", L"deleteRow"},
        {L"rtkG_rowUp", L"rowUp"},             {L"rtkG_rowDown", L"rowDown"},           {L"rtkG_newBar", L"new=bar"},
        {L"rtkG_newMilestone", L"new=milestone"}, {L"rtkG_newBracket", L"new=bracket"}, {L"rtkG_newShade", L"new=shade"},
        {L"rtkG_newPentagon", L"new=pentagon"}, {L"rtkG_newChevron", L"new=chevron"},   {L"rtkG_deleteItem", L"deleteItem"},
        {L"rtkG_fit", L"fit"},                 {L"rtkG_autoScales", L"scales=auto"},    {L"rtkG_calendar", L"calendar"},
        {L"rtkG_datasheet", L"datasheet"},
    };
    const auto it = kIds.find(id);
    return it == kIds.end() ? id : it->second;
}

// Umschalter des PowerPoint-Kontextmenüs → Befehl mit Zustand
std::wstring commandOfToggle(const std::wstring& id, bool pressed) {
    const std::wstring on = pressed ? L"1" : L"0";
    static const std::map<std::wstring, std::wstring> kKeys = {
        {L"rtkT_scaleYear", L"scale:year"},     {L"rtkT_scaleQuarter", L"scale:quarter"}, {L"rtkT_scaleMonth", L"scale:month"},
        {L"rtkT_scaleWeek", L"scale:week"},     {L"rtkT_scaleDay", L"scale:day"},         {L"rtkT_separators", L"separators"},
        {L"rtkT_rowSeparators", L"rowSeparators"}, {L"rtkT_milestoneLabels", L"milestoneLabels"}, {L"rtkT_milestoneLines", L"milestoneLines"},
        {L"rtkT_headlines", L"headlines"},       {L"rtkT_today", L"today"},               {L"rtkT_responsible", L"column:responsible"},
        {L"rtkT_remark", L"column:remark"},      {L"rtkT_rowShade", L"rowShade"},
    };
    if (id == L"rtkT_weekendsShaded") return pressed ? L"weekends=shaded" : L"weekends=plain";
    if (id == L"rtkT_weekendsHidden") return pressed ? L"weekends=hidden" : L"weekends=shaded";
    if (id == L"rtkT_barLabels") return pressed ? L"barLabels=range" : L"barLabels=none";
    const auto it = kKeys.find(id);
    if (it == kKeys.end()) throw Error("Unbekannter Schalter: " + narrow(id));
    return it->second + L"=" + on;
}

std::optional<TimeUnit> unitOf(const std::wstring& name) {
    static const std::pair<const wchar_t*, TimeUnit> kUnits[] = {
        {L"year", TimeUnit::Year}, {L"quarter", TimeUnit::Quarter}, {L"month", TimeUnit::Month}, {L"week", TimeUnit::Week}, {L"day", TimeUnit::Day}};
    for (auto& [n, u] : kUnits)
        if (name == n) return u;
    return std::nullopt;
}

std::wstring& formatOf(GanttOptions& g, TimeUnit unit) {
    switch (unit) {
        case TimeUnit::Day: return g.dayFormat;
        case TimeUnit::Week: return g.weekFormat;
        case TimeUnit::Month: return g.monthFormat;
        case TimeUnit::Quarter: return g.quarterFormat;
        case TimeUnit::Year: return g.yearFormat;
    }
    return g.dayFormat;
}

template <size_t N>
int choice(const std::wstring& value, const wchar_t* const (&names)[N]) {
    for (size_t i = 0; i < N; ++i)
        if (value == names[i]) return static_cast<int>(i);
    throw std::invalid_argument("Unbekannter Wert: " + narrow(value));
}

bool flag(const std::wstring& value) { return value == L"1" || value == L"true" || value == L"on"; }

// Optionen setzen (alles außer Zeilen- und Elementaktionen). shown: gezeigte Skalen (für das erste manuelle Umschalten).
bool setOption(GanttOptions& g, const std::wstring& key, const std::wstring& value, const std::vector<TimeUnit>& shown, const GanttHit& hit) {
    static constexpr const wchar_t* kDesigns[] = {L"standard", L"modern", L"contrast", L"minimal"};
    static constexpr const wchar_t* kWeekends[] = {L"plain", L"shaded", L"hidden"};
    static constexpr const wchar_t* kLabels[] = {L"none", L"range", L"start", L"end"};
    static constexpr const wchar_t* kStyles[] = {L"bracket", L"block", L"text"};
    if (key == L"design") {
        g.design = static_cast<GanttDesign>(choice(value, kDesigns));
        // ein neues Design setzt die von ihm bestimmten Einzelangaben zurück
        g.barHeight.reset(), g.barShape.clear(), g.milestoneShape.clear(), g.categoryStyle.reset(), g.categoryColors.reset();
        g.separators.reset(), g.rowSeparators.reset();
    } else if (key == L"barHeight") {
        g.barHeight = std::clamp(std::stod(value), 0.15, 0.95);
    } else if (key == L"barShape") {
        g.barShape = canonicalBarShape(value);
    } else if (key == L"milestoneShape") {
        g.milestoneShape = canonicalMarker(value);
    } else if (key == L"weekends") {
        g.weekends = static_cast<WeekendMode>(choice(value, kWeekends));
    } else if (key == L"separators") {
        g.separators = flag(value);
    } else if (key == L"rowSeparators") {
        g.rowSeparators = flag(value);
    } else if (key == L"milestoneLines") {
        g.milestoneLines = flag(value);
    } else if (key == L"milestoneLabels") {
        g.milestoneLabels = flag(value);
    } else if (key == L"headlines") {
        g.headlines = flag(value);
    } else if (key == L"today") {
        g.todayLine = flag(value), g.today.clear();
    } else if (key == L"links") {
        g.links = flag(value);
    } else if (key == L"barLabels") {
        g.barLabels = static_cast<BarLabel>(choice(value, kLabels));
    } else if (key == L"dateFormat") {
        g.dateFormat = value;
    } else if (key == L"categoryStyle") {
        g.categoryStyle = static_cast<CategoryStyle>(choice(value, kStyles));
    } else if (key == L"categoryColors") {
        g.categoryColors = flag(value);
    } else if (key == L"column:category") {
        g.category = flag(value);
    } else if (key == L"column:responsible") {
        g.responsible = flag(value);
    } else if (key == L"column:remark") {
        g.remark = flag(value);
    } else if (key == L"scales") {
        if (value != L"auto") throw std::invalid_argument("scales=auto erwartet");
        g.scales.clear();
    } else if (key.starts_with(L"scale:")) {
        const auto unit = unitOf(key.substr(6));
        if (!unit) throw std::invalid_argument("Unbekannte Skala");
        // Skalen von Hand: die automatische Wahl ist danach aus (wie im Original)
        if (g.scales.empty()) g.scales = shown;
        const auto it = std::find(g.scales.begin(), g.scales.end(), *unit);
        if (flag(value) && it == g.scales.end()) g.scales.push_back(*unit);
        if (!flag(value) && it != g.scales.end()) g.scales.erase(it);
        if (g.scales.empty()) throw std::invalid_argument("Mindestens eine Skala muss sichtbar bleiben.");
    } else if (key.starts_with(L"format:")) {
        const auto unit = unitOf(key.substr(7));
        if (!unit) throw std::invalid_argument("Unbekannte Skala");
        formatOf(g, *unit) = value;  // leer = automatisch
    } else if (key == L"rowShade") {
        if (hit.row < gantt::kFirstActivityRow) throw std::invalid_argument("Bitte mit der rechten Maustaste auf eine Zeile klicken.");
        auto& rows = g.shadedRows;
        rows.erase(std::remove(rows.begin(), rows.end(), hit.row), rows.end());
        if (flag(value)) rows.push_back(hit.row);
    } else {
        return false;
    }
    return true;
}

HWND ownerWindow() {
    HWND frame = powerPointFrame();
    return frame ? frame : GetActiveWindow();
}

// Kategorien sortieren: Reihenfolge aus dem Befehl oder aus dem Dialog
void sortCategories(IDispatch* app, IDispatch* element, const std::wstring& value) {
    std::vector<std::wstring> order;
    if (value.empty()) {
        const GanttModel m = ganttModel(element);
        const auto categories = ganttCategories(m.item);
        if (categories.empty()) throw Error("Es gibt noch keine Kategorien. Bitte im Datenblatt in der Spalte „Kategorie“ eintragen.");
        const auto chosen = chooseCategoryOrder(ownerWindow(), categories);
        if (!chosen) return;
        order = *chosen;
    } else {
        for (size_t pos = 0; pos <= value.size();) {
            const size_t sep = std::min(value.find(L'|', pos), value.size());
            if (sep > pos) order.push_back(value.substr(pos, sep - pos));
            pos = sep + 1;
        }
    }
    modifyGantt(app, element, [&](DataItem& item, ChartOptions& o) { sortGanttByCategory(item, o, order); });
}

}  // namespace

void ganttCommand(IDispatch* app, const std::wstring& command, const GanttHit& hit) {
    auto element = selectedGantt(app);
    IDispatch* e = element.Get();
    const std::wstring cmd = commandOfId(command);
    const size_t eq = cmd.find(L'=');
    const std::wstring key = cmd.substr(0, eq), value = eq == std::wstring::npos ? L"" : cmd.substr(eq + 1);

    if (key == L"insertRow") {
        const int before = hit.row >= gantt::kFirstActivityRow ? hit.row + 1 : std::numeric_limits<int>::max();
        modifyGantt(app, e, [&](DataItem& item, ChartOptions& o) { insertGanttRow(item, o, before, L"Neue Aktivität"); });
    } else if (key == L"duplicateRow") {
        const int row = requireRow(hit);
        modifyGantt(app, e, [&](DataItem& item, ChartOptions& o) { duplicateGanttRow(item, o, row, false); });
    } else if (key == L"deleteRow") {
        const int row = requireRow(hit);
        modifyGantt(app, e, [&](DataItem& item, ChartOptions& o) { deleteGanttRow(item, o, row); });
    } else if (key == L"rowUp" || key == L"rowDown") {
        const int row = requireRow(hit);
        modifyGantt(app, e, [&](DataItem& item, ChartOptions& o) {
            if (!moveGanttRow(item, o, row, key == L"rowUp" ? -1 : 1)) throw std::invalid_argument("Die Zeile steht schon am Rand.");
        });
    } else if (key == L"new") {
        static const std::pair<const wchar_t*, GanttItem::Kind> kKinds[] = {
            {L"bar", GanttItem::Kind::Bar},         {L"milestone", GanttItem::Kind::Milestone}, {L"bracket", GanttItem::Kind::Bracket},
            {L"shade", GanttItem::Kind::Shade},     {L"pentagon", GanttItem::Kind::Pentagon},   {L"chevron", GanttItem::Kind::Chevron}};
        const auto it = std::find_if(std::begin(kKinds), std::end(kKinds), [&](auto& k) { return value == k.first; });
        if (it == std::end(kKinds)) throw Error("Unbekannte Elementart: " + narrow(value));
        addItem(app, e, it->second, hit);
    } else if (key == L"deleteItem") {
        if (hit.role.empty()) throw Error("Bitte mit der rechten Maustaste auf einen Balken, Meilenstein, eine Klammer oder Schattierung klicken.");
        const std::string role = hit.role;
        modifyGantt(app, e, [&](DataItem& item, ChartOptions& o) { deleteGanttTarget(item, o, ganttTarget(item, o, role)); });
    } else if (key == L"rowShape" || key == L"rowMarker") {
        // Form des angeklickten Elements: Datenblatt-Balken/-Meilenstein über Spalte 8/9, frei gesetzte Balken über ihre Art
        const bool marker = key == L"rowMarker";
        const std::wstring canonical = marker ? canonicalMarker(value) : canonicalBarShape(value);
        if (!value.empty() && canonical.empty()) throw Error("Unbekannte Form: " + narrow(value));
        const std::string role = hit.role;
        const int row = hit.row;
        modifyGantt(app, e, [&](DataItem& item, ChartOptions& o) {
            const GanttTarget t = ganttTarget(item, o, role);
            if (!marker && t.item >= 0 && t.kind == GanttTarget::Kind::Bar) {
                GanttItem& it = o.gantt.items[t.item];
                it.kind = canonical == L"pentagon" ? GanttItem::Kind::Pentagon : canonical == L"chevron" ? GanttItem::Kind::Chevron : GanttItem::Kind::Bar;
                return;
            }
            if (row < gantt::kFirstActivityRow) throw std::invalid_argument("Bitte mit der rechten Maustaste auf einen Balken oder Meilenstein klicken.");
            // im Datenblatt deutsch, wie man es dort auch eintippen würde
            static const std::map<std::wstring, std::wstring> kNames = {
                {L"bar", L"Balken"},   {L"rounded", L"Abgerundet"}, {L"pentagon", L"Fünfeck"}, {L"chevron", L"Chevron"}, {L"diamond", L"Raute"},
                {L"triangle", L"Dreieck"}, {L"circle", L"Kreis"},   {L"square", L"Quadrat"},   {L"star", L"Stern"}};
            setGanttText(item, row, marker ? gantt::kMarkerCol : gantt::kShapeCol, canonical.empty() ? L"" : kNames.at(canonical));
        });
    } else if (key == L"moveRow") {
        // Zeile an eine andere Stelle ziehen (Overlay): übernimmt dort die Kategorie, Datenblatt und Excel folgen
        const int row = requireRow(hit), target = std::stoi(value);
        modifyGantt(app, e, [&](DataItem& item, ChartOptions& o) { moveGanttRowTo(item, o, row, target); });
    } else if (key == L"link") {
        const int row = requireRow(hit);
        modifyGantt(app, e, [&](DataItem& item, ChartOptions&) { toggleGanttLink(item, row, value); });
    } else if (key == L"fit") {
        modifyGantt(app, e, [&](DataItem& item, ChartOptions& o) {
            if (!fitGanttRange(item, o)) throw std::invalid_argument("Der Zeitplan enthält noch keine Daten.");
        });
    } else if (key == L"calendar") {
        openGanttCalendar(app, hit.inScale || value == L"range" ? std::string() : hit.role);
    } else if (key == L"datasheet") {
        openDatasheet(app, true);
    } else if (key == L"sortCategories") {
        sortCategories(app, e, value);
    } else {
        std::vector<TimeUnit> shown;
        if (key.starts_with(L"scale:")) shown = ganttModel(e).drawing.scales;
        bool known = true;
        modifyGantt(app, e, [&](DataItem&, ChartOptions& o) { known = setOption(o.gantt, key, value, shown, hit); });
        if (!known) throw Error("Unbekannter Befehl: " + narrow(command));
    }
}

GanttState ganttState(IDispatch* app, const GanttHit& hit) {
    auto element = selectedGantt(app);
    const GanttModel m = ganttModel(element.Get());
    GanttState s;
    s.options = m.options.gantt;
    s.scales = m.drawing.scales;
    const GanttLook look = ganttLook(s.options);
    s.barHeight = look.barHeight, s.barShape = look.barShape, s.marker = look.marker;
    s.categoryStyle = look.categoryStyle, s.categoryColors = look.categoryColors, s.separators = look.separators, s.rowSeparators = look.rowSeparators, s.milestoneLines = look.milestoneLines;
    auto drawn = [&](const char* role) { return std::any_of(m.drawing.prims.begin(), m.drawing.prims.end(), [&](const Prim& p) { return p.role == role; }); };
    s.categoryShown = drawn("headline category"), s.responsibleShown = drawn("headline responsible"), s.remarkShown = drawn("headline remark");
    for (size_t r = gantt::kFirstActivityRow; r < m.item.table.size(); ++r)
        s.activities.push_back(m.item.table[r].empty() ? L"" : cellText(m.item.table[r][gantt::kLabelCol]));
    s.categories = ganttCategories(m.item);
    if (!hit.role.empty()) {
        const GanttTarget t = ganttTarget(m.item, m.options, hit.role);
        s.targetIsItem = t.item >= 0;
        switch (t.kind) {
            case GanttTarget::Kind::Bar: s.targetKind = "bar"; break;
            case GanttTarget::Kind::Milestone: s.targetKind = "milestone"; break;
            case GanttTarget::Kind::Bracket: s.targetKind = "bracket"; break;
            case GanttTarget::Kind::Shade: s.targetKind = "shade"; break;
            case GanttTarget::Kind::None: break;
        }
        if (t.item >= 0 && t.kind == GanttTarget::Kind::Bar) {
            const auto kind = m.options.gantt.items[t.item].kind;
            s.rowShape = kind == GanttItem::Kind::Pentagon ? L"pentagon" : kind == GanttItem::Kind::Chevron ? L"chevron" : L"";
        }
    }
    if (hit.row >= gantt::kFirstActivityRow && static_cast<size_t>(hit.row) < m.item.table.size()) {
        const Row& row = m.item.table[hit.row];
        auto text = [&](int c) { return static_cast<size_t>(c) < row.size() ? cellText(row[c]) : std::wstring(); };
        if (!s.targetIsItem) s.rowShape = canonicalBarShape(text(gantt::kShapeCol));
        s.rowMarker = canonicalMarker(text(gantt::kMarkerCol));
        const std::wstring links = text(gantt::kPredecessorCol);
        for (size_t pos = 0; pos <= links.size();) {
            const size_t sep = std::min(links.find_first_of(L";,", pos), links.size());
            std::wstring name = links.substr(pos, sep - pos);
            const auto a = name.find_first_not_of(L" \t"), b = name.find_last_not_of(L" \t");
            if (a != std::wstring::npos) s.rowLinks.push_back(name.substr(a, b - a + 1));
            pos = sep + 1;
        }
        s.rowShaded = std::find(s.options.shadedRows.begin(), s.options.shadedRows.end(), hit.row) != s.options.shadedRows.end();
    }
    s.today = m.options.gantt.referenceDay;
    s.language = m.options.language;
    return s;
}

bool ganttPressed(IDispatch* app, const std::wstring& id, const GanttHit& hit) {
    const GanttState s = ganttState(app, hit);
    const GanttOptions& g = s.options;
    auto scaleShown = [&](TimeUnit u) { return std::find(s.scales.begin(), s.scales.end(), u) != s.scales.end(); };
    if (id == L"rtkT_scaleYear") return scaleShown(TimeUnit::Year);
    if (id == L"rtkT_scaleQuarter") return scaleShown(TimeUnit::Quarter);
    if (id == L"rtkT_scaleMonth") return scaleShown(TimeUnit::Month);
    if (id == L"rtkT_scaleWeek") return scaleShown(TimeUnit::Week);
    if (id == L"rtkT_scaleDay") return scaleShown(TimeUnit::Day);
    if (id == L"rtkT_weekendsShaded") return g.weekends == WeekendMode::Shaded;
    if (id == L"rtkT_weekendsHidden") return g.weekends == WeekendMode::Hidden;
    if (id == L"rtkT_separators") return s.separators;
    if (id == L"rtkT_rowSeparators") return s.rowSeparators;
    if (id == L"rtkT_barLabels") return g.barLabels != BarLabel::None;
    if (id == L"rtkT_milestoneLabels") return g.milestoneLabels;
    if (id == L"rtkT_milestoneLines") return s.milestoneLines;
    if (id == L"rtkT_headlines") return g.headlines;
    if (id == L"rtkT_today") return g.todayLine;
    if (id == L"rtkT_rowShade") return s.rowShaded;
    if (id == L"rtkT_responsible") return s.responsibleShown;
    if (id == L"rtkT_remark") return s.remarkShown;
    return false;
}

void ganttToggle(IDispatch* app, const std::wstring& id, bool pressed, const GanttHit& hit) { ganttCommand(app, commandOfToggle(id, pressed), hit); }

void setGanttDates(IDispatch* app, const std::string& role, const std::wstring& first, const std::wstring& last) {
    auto element = selectedGantt(app);
    const auto a = parseIsoDate(first), b = parseIsoDate(last.empty() ? first : last);
    if (!a || !b || *b < *a) throw Error("Datum als ISO angeben, z. B. 2026-11-02 (Ende nicht vor dem Anfang)");
    modifyGantt(app, element.Get(), [&](DataItem& item, ChartOptions& o) {
        if (role.empty() || role == "range") {
            setGanttRange(o, *a, *b + 1);
            return;
        }
        const GanttTarget t = ganttTarget(item, o, role);
        if (t.kind == GanttTarget::Kind::None) throw std::invalid_argument("Dieses Element gibt es nicht.");
        engine::setGanttDates(item, o, t, *a, t.kind == GanttTarget::Kind::Milestone ? *a : *b + 1);
    });
}

void openGanttCalendar(IDispatch* app, const std::string& role) {
    auto element = selectedGantt(app);
    const GanttModel m = ganttModel(element.Get());
    double first = m.drawing.axis.t0, last = m.drawing.axis.t1 - 1;
    bool single = false;
    std::wstring title = L"rethink – Datumsbereich";
    if (!role.empty() && role != "range") {
        const GanttTarget t = ganttTarget(m.item, m.options, role);
        if (t.kind == GanttTarget::Kind::None) throw Error("Dieses Element hat keine Daten.");
        single = t.kind == GanttTarget::Kind::Milestone;
        first = std::floor(t.start);
        last = single ? first : std::ceil(t.end - 1e-9) - 1;
        title = single ? L"rethink – Meilenstein" : L"rethink – Zeitraum";
    }
    ComPtr<IDispatch> appRef = app;
    const long slideId = disp::call(disp::object(element.Get(), L"Parent").Get(), L"SlideID").toLong();
    const std::wstring name = disp::call(element.Get(), L"Name").str();
    CalendarCallbacks callbacks;
    callbacks.changed = [appRef, role, slideId, name](double a, double b) {
        try {
            // Das Element über Folie und Namen wiederfinden: nach einem vollständigen Neuzeichnen ist es eine neue Form
            auto slide = disp::object(disp::object(disp::object(appRef.Get(), L"ActivePresentation").Get(), L"Slides").Get(), L"FindBySlideID", slideId);
            auto group = disp::object(disp::object(slide.Get(), L"Shapes").Get(), L"Item", name);
            disp::call(group.Get(), L"Select");
            setGanttDates(appRef.Get(), role, isoDate(a), isoDate(b));
        } catch (const std::exception& e) {
            MessageBoxW(nullptr, widen(e.what()).c_str(), kTitle, MB_ICONWARNING);
        }
    };
    openCalendarWindow(powerPointFrame(), title, first, last, single, std::move(callbacks), screenRectOf(app, element.Get()));
}

std::wstring ganttTargets(IDispatch* app) {
    auto element = selectedGantt(app);
    const GanttModel m = ganttModel(element.Get());
    std::wstring text = std::format(L"range={}..{}", isoDate(m.drawing.axis.t0), isoDate(m.drawing.axis.t1 - 1));
    std::wstring scales;
    for (auto u : m.drawing.scales) scales += (scales.empty() ? L"" : L",") + unitName(u, L"en");
    text += L"; scales=" + scales;
    for (auto& p : m.drawing.prims) {
        if (!draggable(p.flags)) continue;
        const GanttTarget t = ganttTarget(m.item, m.options, p.role);
        if (t.kind == GanttTarget::Kind::None) continue;
        text += std::format(L"; {}={}", widen(p.role), isoDate(t.start));
        if (t.kind != GanttTarget::Kind::Milestone) text += L".." + isoDate(t.end);
    }
    return text;
}

void selectInnerShape(IDispatch* element, const std::string& role) {
    disp::call(disp::object(disp::object(element, L"GroupItems").Get(), L"Item", widen(role)).Get(), L"Select");
}

std::vector<SegmentInfo> segmentsOf(IDispatch* element) {
    std::vector<SegmentInfo> result;
    auto items = disp::object(element, L"GroupItems");
    const long count = disp::call(items.Get(), L"Count").toLong();
    for (long i = 1; i <= count; ++i) {
        auto shape = disp::object(items.Get(), L"Item", i);
        const std::wstring name = disp::call(shape.Get(), L"Name").str();
        SegmentInfo s;
        if (swscanf_s(name.c_str(), L"segment r%d c%d", &s.row, &s.column) != 2) continue;
        s.flags = narrow(tagValue(shape.Get(), kSegmentTag));
        const Rect r = rectOf(shape.Get());
        s.left = r.left, s.top = r.top, s.width = r.width, s.height = r.height;
        result.push_back(s);
    }
    return result;
}

bool relayoutIfResized(IDispatch* element) {
    const auto stored = storedLayout(element);
    if (!stored) return false;
    const Rect now = rectOf(element);
    if (std::fabs(now.width - stored->group.width) < 0.75 && std::fabs(now.height - stored->group.height) < 0.75) return false;
    const std::wstring model = tagValue(element, kModelTag);
    const DataItem item = onWorker([&] { return parseDataItemJson(model); });
    renderElement(disp::object(element, L"Parent").Get(), element, item, layoutBox(element));
    return true;
}

int watchCount() { return static_cast<int>(g_sheets.size()); }

void stopDatasheetWatch() {
    std::vector<Datasheet*> grids;
    for (auto& [key, g] : g_grids) grids.push_back(g->window);
    for (auto* w : grids) datasheetClose(w);
    g_gridApp.Reset();
    if (g_timer) KillTimer(nullptr, g_timer);
    g_timer = 0;
    for (auto& [key, s] : g_sheets) {
        if (s.workbook) {
            try {
                disp::put(s.workbook.Get(), L"Saved", true);
                disp::call(s.workbook.Get(), L"Close", false);
            } catch (...) {
            }
        }
        releaseSheet(s);
    }
    g_sheets.clear();
    quitExcelIfUnused();
}

void commandInsertWaterfall(IDispatch* app) { commandInsertElement(app, ElementKind::Waterfall, Orientation::Up); }

void commandInsertElement(IDispatch* app, ElementKind kind, Orientation orientation) {
    try {
        insertElement(app, kind, orientation, false);  // Datenblatt erst auf Wunsch (Symbol unten rechts, Doppelklick)
    } catch (const std::exception& e) {
        MessageBoxW(nullptr, widen(e.what()).c_str(), kTitle, MB_ICONERROR);
    } catch (const winrt::hresult_error& e) {
        MessageBoxW(nullptr, e.message().c_str(), kTitle, MB_ICONERROR);
    }
}

void commandOpenDatasheet(IDispatch* app) {
    try {
        openDatasheet(app, true);
    } catch (const std::exception& e) {
        MessageBoxW(nullptr, widen(e.what()).c_str(), kTitle, MB_ICONERROR);
    } catch (const winrt::hresult_error& e) {
        MessageBoxW(nullptr, e.message().c_str(), kTitle, MB_ICONERROR);
    }
}

}  // namespace rethink::addin
