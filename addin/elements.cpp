#include <winrt/base.h>

#include "elements.h"

#include "datasheet.h"
#include "datasheet_window.h"
#include "dispatch.h"
#include "engine/chart.h"
#include "ppttc.h"
#include "render_element.h"
#include "xlsx.h"

#include <ocidl.h>
#include <shellapi.h>

#include <cmath>
#include <exception>
#include <filesystem>
#include <format>
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
constexpr wchar_t kSegmentTag[] = L"RETHINK_SEG";    // "pos", "neg", "e"
constexpr long kMsoShapeRectangle = 1;
constexpr long kMsoTextOrientationHorizontal = 1;
constexpr long kMsoGroup = 6;
constexpr long kMsoAnchorMiddle = 3;
constexpr long kMsoLineDash = 4;
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
        case Color::Kind::Scheme: return L"s:" + c.scheme;
        case Color::Kind::Rgb: return std::format(L"r:{:06X}", c.rgb);
        case Color::Kind::None: break;
    }
    return L"n";
}

void applyColor(IDispatch* colorFormat, const Color& c) {
    if (c.kind == Color::Kind::Scheme)
        disp::put(colorFormat, L"ObjectThemeColor", themeColor(c.scheme));
    else if (c.kind == Color::Kind::Rgb)
        disp::put(colorFormat, L"RGB", static_cast<long>(((c.rgb >> 16) & 0xFF) | (c.rgb & 0xFF00) | ((c.rgb & 0xFF) << 16)));
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

// Eine Grundform über das Objektmodell zeichnen.
ComPtr<IDispatch> drawPrim(IDispatch* shapes, const Prim& p) {
    switch (p.kind) {
        case Prim::Kind::Rect: {
            auto shape = disp::object(shapes, L"AddShape", kMsoShapeRectangle, pt(p.x), pt(p.y), pt(p.w), pt(p.h));
            auto fill = disp::object(shape.Get(), L"Fill");
            disp::call(fill.Get(), L"Solid");
            applyColor(disp::object(fill.Get(), L"ForeColor").Get(), p.fill);
            disp::put(disp::object(shape.Get(), L"Line").Get(), L"Visible", 0L);
            disp::put(disp::object(shape.Get(), L"Shadow").Get(), L"Visible", 0L);
            return shape;
        }
        case Prim::Kind::Line: {
            auto shape = disp::object(shapes, L"AddLine", pt(p.x), pt(p.y), pt(p.x + p.w), pt(p.y + p.h));
            auto line = disp::object(shape.Get(), L"Line");
            disp::put(line.Get(), L"Weight", p.strokePt);
            applyColor(disp::object(line.Get(), L"ForeColor").Get(), p.stroke);
            if (p.dashed) disp::put(line.Get(), L"DashStyle", kMsoLineDash);
            return shape;
        }
        case Prim::Kind::Text: {
            auto shape = disp::object(shapes, L"AddTextbox", kMsoTextOrientationHorizontal, pt(p.x), pt(p.y), pt(p.w), pt(p.h));
            auto frame = disp::object(shape.Get(), L"TextFrame");
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
            applyColor(disp::object(font.Get(), L"Color").Get(), p.textColor);
            disp::put(disp::object(range.Get(), L"ParagraphFormat").Get(), L"Alignment",
                      p.align == Align::Left ? 1L : p.align == Align::Right ? 3L : 2L);
            setGeometry(shape.Get(), p);
            return shape;
        }
    }
    throw Error("Unbekannte Grundform");
}

// Rolle, Farbsignatur und Segmentart an einer Form vermerken.
void markShape(IDispatch* shape, const Prim& p) {
    disp::put(shape, L"Name", widen(p.role));
    if (p.kind == Prim::Kind::Rect) setTag(shape, kStyleTag, signature(p.fill));
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
    setGeometry(shape, p);
    if (p.kind == Prim::Kind::Rect) {
        const auto sig = signature(p.fill);
        if (tagValue(shape, kStyleTag) != sig) {
            applyColor(disp::object(disp::object(shape, L"Fill").Get(), L"ForeColor").Get(), p.fill);
            setTag(shape, kStyleTag, sig);
        }
        if (tagValue(shape, kSegmentTag) != widen(p.flags)) setTag(shape, kSegmentTag, widen(p.flags));
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
};

Prepared prepare(const DataItem& item, Box bounds) {
    const std::wstring separator = localeDecimalSeparator();
    return onWorker([&] {
        ChartOptions options = parseChartOptions(item.chartJson);
        options.decimalSeparator = separator;
        Prepared p;
        try {
            p.drawing = layoutChart(item, options, bounds);
        } catch (const std::invalid_argument& e) {
            throw Error(std::format("Diagramm '{}': {}", narrow(item.name), e.what()));
        }
        p.model = serializeDataItem(item);
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

// Zeichnet ein Element (neu oder an Ort und Stelle) und vermerkt Modell und Layout.
ComPtr<IDispatch> renderElement(IDispatch* slide, ComPtr<IDispatch> group, const DataItem& item, Box box) {
    const Prepared prepared = prepare(item, box);
    if (!group || !updateInPlace(group.Get(), prepared.drawing)) {
        long zorder = 0;
        if (group) {
            zorder = disp::call(group.Get(), L"ZOrderPosition").toLong();
            disp::call(group.Get(), L"Delete");
        }
        group = drawElement(disp::object(slide, L"Shapes").Get(), prepared.drawing, item.name);
        if (zorder > 0)
            for (int i = 0; i < 500 && disp::call(group.Get(), L"ZOrderPosition").toLong() > zorder; ++i)
                disp::call(group.Get(), L"ZOrder", 3L);  // msoSendBackward: alte Stapelposition
    }
    setTag(group.Get(), kModelTag, prepared.model);
    storeLayout(group.Get(), prepared.drawing.bounds);
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

DataItem sampleWaterfall(const std::wstring& name) {
    auto text = [](const wchar_t* s) { return Cell{CellKind::String, s, 0, std::nullopt}; };
    auto num = [](double v) { return Cell{CellKind::Number, L"", v, std::nullopt}; };
    DataItem item;
    item.name = name;
    item.chartJson = LR"({"type":"waterfall","datatable":true})";
    item.table = {
        {Cell{}, text(L"Start"), text(L"Umsatz"), text(L"Kosten"), text(L"Zwischensumme"), text(L"Sonstiges"), text(L"Ende")},
        {text(L"Wert"), num(100), num(40), num(-25), text(L"e"), num(-10), text(L"e")},
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

std::wstring insertWaterfall(IDispatch* app, bool openSheet) {
    auto slide = currentSlide(app);
    auto pageSetup = disp::object(disp::object(activeWindow(app).Get(), L"Presentation").Get(), L"PageSetup");
    const double slideWidth = disp::call(pageSetup.Get(), L"SlideWidth").toDouble();
    const double slideHeight = disp::call(pageSetup.Get(), L"SlideHeight").toDouble();
    const double width = slideWidth * 0.7, height = slideHeight * 0.6;

    const std::wstring name = freeName(slide.Get(), L"Wasserfall");
    const Box box{emu((slideWidth - width) / 2), emu((slideHeight - height) / 2 + slideHeight * 0.05), emu(width), emu(height)};
    auto group = renderElement(slide.Get(), nullptr, sampleWaterfall(name), box);
    try {
        disp::call(group.Get(), L"Select");
    } catch (const Error&) {
    }
    if (openSheet) openDatasheet(app, true);
    return name;
}

// ---- Eingebautes Datenblatt ----

namespace {

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
    callbacks.closed = [key, ctx]() {
        g_grids.erase(key);
        delete ctx;
    };
    ctx->window = openDatasheetWindow(powerPointFrame(), L"rethink – Datenblatt " + ctx->name, std::move(table), localeDecimalSeparator(),
                                      std::move(callbacks), screenRectOf(app, element.Get()), show);
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
    const std::wstring separator = localeDecimalSeparator();

    SegmentDrag drag;
    drag.slide = disp::object(element.Get(), L"Parent");
    drag.group = element;
    drag.row = row, drag.column = column;
    drag.box = layoutBox(element.Get());
    std::tie(drag.item, drag.options) = onWorker([&] {
        DataItem item = parseDataItemJson(stored);
        ChartOptions options = parseChartOptions(item.chartJson);
        return std::pair{std::move(item), std::move(options)};
    });
    drag.options.decimalSeparator = separator;
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

void commandInsertWaterfall(IDispatch* app) {
    try {
        insertWaterfall(app, false);  // Datenblatt erst auf Wunsch (Symbol unten rechts, Doppelklick)
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
