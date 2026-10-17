#include <winrt/base.h>

#include "elements.h"

#include "datasheet.h"
#include "dispatch.h"
#include "engine/chart.h"
#include "ppttc.h"
#include "render_element.h"

#include <shellapi.h>

#include <chrono>
#include <exception>
#include <filesystem>
#include <format>
#include <map>
#include <set>
#include <thread>

namespace rethink::addin {
namespace fs = std::filesystem;
using namespace engine;

namespace {

constexpr wchar_t kTitle[] = L"rethink";
constexpr wchar_t kPrefix[] = L"rethink:";
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

void applyColor(IDispatch* colorFormat, const Color& c) {
    if (c.kind == Color::Kind::Scheme)
        disp::put(colorFormat, L"ObjectThemeColor", themeColor(c.scheme));
    else if (c.kind == Color::Kind::Rgb)
        disp::put(colorFormat, L"RGB", static_cast<long>(((c.rgb >> 16) & 0xFF) | (c.rgb & 0xFF00) | ((c.rgb & 0xFF) << 16)));
}

// Eine Grundform über das Objektmodell zeichnen; liefert die neue Form.
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
            disp::put(shape.Get(), L"Left", pt(p.x));
            disp::put(shape.Get(), L"Top", pt(p.y));
            disp::put(shape.Get(), L"Width", pt(p.w));
            disp::put(shape.Get(), L"Height", pt(p.h));
            return shape;
        }
    }
    throw Error("Unbekannte Grundform");
}

// Zeichnet alle Grundformen, gruppiert sie und versieht die Gruppe mit Name und Modell.
ComPtr<IDispatch> drawElement(IDispatch* shapes, const Drawing& d, const std::wstring& name, const std::wstring& model) {
    static unsigned long serial = GetTickCount();
    const std::wstring stem = std::format(L"rethink {:08X}", ++serial);
    std::vector<std::wstring> names;
    for (size_t i = 0; i < d.prims.size(); ++i) {
        auto shape = drawPrim(shapes, d.prims[i]);
        names.push_back(std::format(L"{} {}", stem, i + 1));
        disp::put(shape.Get(), L"Name", names.back());
    }

    // Shapes.Range(Array(...)).Group()
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

    disp::put(group.Get(), L"Name", name);
    disp::put(group.Get(), L"AlternativeText", std::wstring(kPrefix) + name);
    disp::call(disp::object(group.Get(), L"Tags").Get(), L"Add", std::wstring(kModelTag), model);
    return group;
}

std::wstring tagValue(IDispatch* shape, const wchar_t* tag) {
    try {
        return disp::call(disp::object(shape, L"Tags").Get(), L"Item", tag).str();
    } catch (const Error&) {
        return {};
    }
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

// Ausgewähltes rethink-Element des aktiven Fensters.
ComPtr<IDispatch> selectedElement(IDispatch* app) {
    auto element = rethinkElementOf(disp::object(activeWindow(app).Get(), L"Selection").Get());
    if (!element) throw Error("Bitte zuerst ein rethink-Diagramm auswählen (Wasserfall einfügen oder per .ppttc erzeugen).");
    return element;
}

Box boundsOf(IDispatch* shape) {
    return {emu(disp::call(shape, L"Left").toDouble()), emu(disp::call(shape, L"Top").toDouble()), emu(disp::call(shape, L"Width").toDouble()),
            emu(disp::call(shape, L"Height").toDouble())};
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

// ---- Datenblatt-Überwachung ----

struct Watch {
    fs::path file;
    fs::file_time_type stamp;
    ComPtr<IDispatch> slide;
    long shapeId = 0;
    std::wstring name;
};

std::map<std::wstring, Watch> g_watches;  // Schlüssel: Datei (klein)
ComPtr<IDispatch> g_app;
UINT_PTR g_timer = 0;
bool g_busy = false;

ComPtr<IDispatch> findShape(IDispatch* slide, long id) {
    auto shapes = disp::object(slide, L"Shapes");
    const long count = disp::call(shapes.Get(), L"Count").toLong();
    for (long i = 1; i <= count; ++i) {
        auto shape = disp::object(shapes.Get(), L"Item", i);
        if (disp::call(shape.Get(), L"Id").toLong() == id) return shape;
    }
    return nullptr;
}

// Zeichnet das Element einer Überwachung aus dem Datenblatt neu. Gibt false zurück, wenn sich nichts geändert hat.
bool refresh(Watch& w) {
    std::error_code ec;
    const auto stamp = fs::last_write_time(w.file, ec);
    if (ec) throw Error("Das Datenblatt wurde gelöscht.");
    if (stamp == w.stamp) return false;

    auto group = findShape(w.slide.Get(), w.shapeId);
    if (!group) throw Error(std::format("Das Diagramm '{}' ist nicht mehr auf der Folie.", narrow(w.name)));
    const std::wstring stored = tagValue(group.Get(), kModelTag);
    const auto bytes = readFileShared(w.file);
    if (bytes.empty()) return false;  // Excel schreibt gerade

    DataItem item = onWorker([&] {
        DataItem parsed = parseDataItemJson(stored);
        parsed.table = tableFromDatasheet(bytes);
        return parsed;
    });
    const Prepared prepared = prepare(item, boundsOf(group.Get()));
    w.stamp = stamp;

    const long zorder = disp::call(group.Get(), L"ZOrderPosition").toLong();
    auto shapes = disp::object(w.slide.Get(), L"Shapes");
    const std::wstring name = disp::call(group.Get(), L"Name").str();
    disp::call(group.Get(), L"Delete");
    auto fresh = drawElement(shapes.Get(), prepared.drawing, name, prepared.model);
    for (int i = 0; i < 500 && disp::call(fresh.Get(), L"ZOrderPosition").toLong() > zorder; ++i)
        disp::call(fresh.Get(), L"ZOrder", 3L);  // msoSendBackward: alte Stapelposition wiederherstellen
    w.shapeId = disp::call(fresh.Get(), L"Id").toLong();
    return true;
}

int refreshAll(bool reportErrors) {
    if (g_busy) return 0;
    g_busy = true;
    int redrawn = 0;
    std::vector<std::wstring> finished;
    for (auto& [key, w] : g_watches) {
        try {
            if (refresh(w)) ++redrawn;
        } catch (const std::exception& e) {
            const std::wstring text = widen(e.what());
            std::error_code ec;
            const bool gone = !fs::exists(w.file, ec);
            if (gone || text.find(L"nicht mehr auf der Folie") != std::wstring::npos) {
                finished.push_back(key);
            } else {
                w.stamp = fs::last_write_time(w.file, ec);  // dieselbe Fassung nicht erneut melden
            }
            if (reportErrors && !gone) MessageBoxW(nullptr, (L"Datenblatt „" + w.name + L"“: " + text).c_str(), kTitle, MB_ICONWARNING);
        } catch (...) {
            finished.push_back(key);  // Folie/Präsentation nicht mehr erreichbar
        }
    }
    for (auto& key : finished) g_watches.erase(key);
    if (g_watches.empty() && g_timer) {
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

// Datenblätter immer mit Excel öffnen (die .xlsx-Zuordnung kann auf ein anderes Programm zeigen);
// nur wenn Excel fehlt, die Standardzuordnung.
bool openInExcel(const fs::path& file) {
    const std::wstring quoted = L"\"" + file.wstring() + L"\"";
    if (reinterpret_cast<INT_PTR>(ShellExecuteW(nullptr, L"open", L"excel.exe", quoted.c_str(), nullptr, SW_SHOWNORMAL)) > 32) return true;
    return reinterpret_cast<INT_PTR>(ShellExecuteW(nullptr, L"open", file.c_str(), nullptr, nullptr, SW_SHOWNORMAL)) > 32;
}

}  // namespace

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

std::wstring elementLabel(IDispatch* element) {
    const std::wstring model = tagValue(element, kModelTag);
    const std::wstring kind = model.find(L"\"waterfall\"") != std::wstring::npos ? L"Wasserfall" : L"Diagramm";
    return std::format(L"rethink · {} · {}", kind, disp::call(element, L"Name").str());
}

std::wstring insertWaterfall(IDispatch* app, bool openSheet) {
    auto slide = currentSlide(app);
    auto pageSetup = disp::object(disp::object(activeWindow(app).Get(), L"Presentation").Get(), L"PageSetup");
    const double slideWidth = disp::call(pageSetup.Get(), L"SlideWidth").toDouble();
    const double slideHeight = disp::call(pageSetup.Get(), L"SlideHeight").toDouble();
    const double width = slideWidth * 0.7, height = slideHeight * 0.6;

    const std::wstring name = freeName(slide.Get(), L"Wasserfall");
    const DataItem item = sampleWaterfall(name);
    const Prepared prepared = prepare(item, Box{emu((slideWidth - width) / 2), emu((slideHeight - height) / 2 + slideHeight * 0.05), emu(width), emu(height)});
    auto group = drawElement(disp::object(slide.Get(), L"Shapes").Get(), prepared.drawing, name, prepared.model);
    try {
        disp::call(group.Get(), L"Select");
    } catch (const Error&) {
    }
    if (openSheet) openDatasheet(app, true);
    return name;
}

std::wstring openDatasheet(IDispatch* app, bool show) {
    auto element = selectedElement(app);
    auto slide = disp::object(element.Get(), L"Parent");
    const long slideId = disp::call(slide.Get(), L"SlideID").toLong();
    const long shapeId = disp::call(element.Get(), L"Id").toLong();
    const std::wstring name = disp::call(element.Get(), L"Name").str();
    const std::wstring stored = tagValue(element.Get(), kModelTag);

    const fs::path file = datasheetPath(name, slideId);
    const std::wstring key = toLower(file.wstring());
    if (auto existing = g_watches.find(key); existing != g_watches.end()) {
        existing->second.shapeId = shapeId;  // bereits überwacht: dieselbe Datei weiterverwenden
        existing->second.slide = slide;
    } else {
        const auto bytes = onWorker([&] { return datasheetFromTable(parseDataItemJson(stored).table); });
        HANDLE h = CreateFileW(file.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (h == INVALID_HANDLE_VALUE)
            throw Error("Das Datenblatt ist noch in Excel geöffnet. Bitte dort weiterarbeiten oder die Datei schließen: " + narrow(file.wstring()));
        DWORD written = 0;
        WriteFile(h, bytes.data(), static_cast<DWORD>(bytes.size()), &written, nullptr);
        CloseHandle(h);
        g_watches[key] = Watch{file, fs::last_write_time(file), slide, shapeId, name};
    }
    g_app = app;
    if (!g_timer) g_timer = SetTimer(nullptr, 0, 1000, onTimer);

    if (show && !openInExcel(file)) throw Error("Excel konnte das Datenblatt nicht öffnen: " + narrow(file.wstring()));
    return file.wstring();
}

int watchCount() { return static_cast<int>(g_watches.size()); }

int applyDatasheets(IDispatch*) { return refreshAll(false); }

void stopDatasheetWatch() {
    if (g_timer) KillTimer(nullptr, g_timer);
    g_timer = 0;
    g_watches.clear();
    g_app.Reset();
}

void commandInsertWaterfall(IDispatch* app) {
    try {
        insertWaterfall(app, true);
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
