// COM-Add-in für PowerPoint: IDTExtensibility2 (Laden/Entladen) und IRibbonExtensibility (Ribbon-XML).
// Ribbon-Rückrufe und die Automationsschnittstelle laufen über dieselbe IDispatch-Implementierung.

#include <winrt/base.h>

#include "calendar_window.h"
#include "commands.h"
#include "dispatch.h"
#include "elements.h"
#include "engine/dates.h"
#include "events.h"
#include "gallery.h"
#include "gantt_menu.h"
#include "officeapi.h"
#include "resource.h"
#include "version.h"

#include <wrl/implements.h>
#include <wrl/module.h>

// WRL (module.h, StaticStorage) löst bei der Instanziierung C4324 aus: Auffüllung wegen alignas – unkritisch.
#pragma warning(disable : 4324)

#include <algorithm>
#include <format>
#include <string>
#include <vector>

extern "C" IMAGE_DOS_HEADER __ImageBase;

namespace rethink::addin {
namespace {

using Microsoft::WRL::ChainInterfaces;
using Microsoft::WRL::ClassicCom;
using Microsoft::WRL::RuntimeClass;
using Microsoft::WRL::RuntimeClassFlags;

enum : DISPID {
    // Ribbon-Rückrufe (onAction im Ribbon-XML)
    kOnBuild = 1,
    kOnInsertChart,
    kOnSetName,
    kOnInsertWaterfall,
    kOnOpenDatasheet,
    kOnInsertGallery,
    kOnInsertGantt,
    kOnGanttVisible,  // getVisible des Kontextmenüs „rethink-Zeitplan“
    kOnGanttAction,   // Schaltflächen darin
    kOnGanttToggle,   // Umschalter: onAction(control, pressed)
    kOnGanttPressed,  // Umschalter: getPressed
    kOnRibbonLoad,    // customUI onLoad: Office hat das Ribbon-XML angenommen
    // Automation über Application.COMAddIns("rethink.addin").Object
    kBuildFrom = 101,
    kInsertChart,
    kSetName,
    kGetName,
    kLastWarnings,
    kVersion,
    kInsertWaterfall,
    kOpenDatasheet,
    kApplyDatasheets,
    kEventStatus,
    kAdjustSegment,
    kOverlayInfo,
    kDatasheetWorkbook,
    kOpenExcelDatasheet,
    kDatasheetEnter,
    kDatasheetPaste,
    kDatasheetCell,
    kCloseDatasheet,
    kDatasheetCount,
    kDragSegment,
    kInsertElement,
    kGanttTargets,
    kGanttDrag,
    kGanttCommand,
    kGanttToggle,
    kGanttPressed,
    kSetGanttDates,
    kGanttHit,
    kOpenCalendar,
    kCloseCalendar,
    kRibbonStatus,
    kActivateTab,
    kGanttMenu,
};

struct Member {
    const wchar_t* name;
    DISPID id;
};

constexpr Member kMembers[] = {
    {L"RethinkBuild", kOnBuild},     {L"RethinkInsertChart", kOnInsertChart}, {L"RethinkSetName", kOnSetName},
    {L"BuildFrom", kBuildFrom},      {L"InsertChart", kInsertChart},          {L"SetName", kSetName},
    {L"GetName", kGetName},          {L"LastWarnings", kLastWarnings},        {L"Version", kVersion},
    {L"RethinkInsertWaterfall", kOnInsertWaterfall}, {L"RethinkOpenDatasheet", kOnOpenDatasheet},
    {L"InsertWaterfall", kInsertWaterfall}, {L"OpenDatasheet", kOpenDatasheet}, {L"ApplyDatasheets", kApplyDatasheets},
    {L"EventStatus", kEventStatus},  {L"AdjustSegment", kAdjustSegment}, {L"OverlayInfo", kOverlayInfo},
    {L"DatasheetWorkbook", kDatasheetWorkbook}, {L"OpenExcelDatasheet", kOpenExcelDatasheet}, {L"DatasheetEnter", kDatasheetEnter},
    {L"DatasheetPaste", kDatasheetPaste}, {L"DatasheetCell", kDatasheetCell}, {L"CloseDatasheet", kCloseDatasheet},
    {L"DatasheetCount", kDatasheetCount}, {L"DragSegment", kDragSegment},
    {L"RethinkInsertGallery", kOnInsertGallery}, {L"RethinkInsertGantt", kOnInsertGantt}, {L"RethinkGanttVisible", kOnGanttVisible},
    {L"RethinkGanttAction", kOnGanttAction}, {L"RethinkGanttToggle", kOnGanttToggle}, {L"RethinkGanttPressed", kOnGanttPressed},
    {L"InsertElement", kInsertElement}, {L"GanttTargets", kGanttTargets}, {L"GanttDrag", kGanttDrag}, {L"GanttCommand", kGanttCommand},
    {L"GanttToggle", kGanttToggle}, {L"GanttPressed", kGanttPressed}, {L"SetGanttDates", kSetGanttDates}, {L"GanttHit", kGanttHit},
    {L"OpenCalendar", kOpenCalendar}, {L"CloseCalendar", kCloseCalendar}, {L"RethinkRibbonLoad", kOnRibbonLoad},
    {L"RibbonStatus", kRibbonStatus}, {L"ActivateTab", kActivateTab}, {L"GanttMenu", kGanttMenu},
};

// Kontextmenü „rethink-Zeitplan“ (im Original: Kontextmenüs von Zeilen, Elementen und Skalen). Steuerelement-IDs
// müssen im ganzen Ribbon eindeutig sein; das Menü erscheint in mehreren PowerPoint-Kontextmenüs, daher je ein Suffix.
std::wstring ganttMenuXml(const std::wstring& suffix) {
    auto button = [&](const wchar_t* id, const wchar_t* label) {
        return std::format(L"<button id=\"{}_{}\" label=\"{}\" onAction=\"RethinkGanttAction\"/>", id, suffix, label);
    };
    auto toggle = [&](const wchar_t* id, const wchar_t* label) {
        return std::format(L"<toggleButton id=\"{}_{}\" label=\"{}\" getPressed=\"RethinkGanttPressed\" onAction=\"RethinkGanttToggle\"/>", id, suffix, label);
    };
    auto separator = [&](int n) { return std::format(L"<menuSeparator id=\"rtkSep{}_{}\"/>", n, suffix); };
    std::wstring xml = std::format(L"<menu id=\"rtkGantt_{}\" label=\"rethink-Zeitplan\" imageMso=\"DateAndTimeInsert\" getVisible=\"RethinkGanttVisible\">",
                                   suffix);
    xml += button(L"rtkG_insertRow", L"Zeile einfügen") + button(L"rtkG_duplicateRow", L"Zeile duplizieren") +
           button(L"rtkG_deleteRow", L"Zeile löschen") + button(L"rtkG_rowUp", L"Zeile nach oben") + button(L"rtkG_rowDown", L"Zeile nach unten") +
           toggle(L"rtkT_rowShade", L"Zeile schattieren");
    xml += separator(1);
    xml += button(L"rtkG_newBar", L"Neuer Balken") + button(L"rtkG_newMilestone", L"Neuer Meilenstein") +
           button(L"rtkG_newBracket", L"Neue Klammer") + button(L"rtkG_newShade", L"Neue Schattierung") +
           button(L"rtkG_newPentagon", L"Neues Fünfeck") + button(L"rtkG_newChevron", L"Neues Chevron") +
           button(L"rtkG_deleteItem", L"Element löschen");
    xml += separator(2);
    xml += button(L"rtkG_calendar", L"Kalender öffnen …") + button(L"rtkG_fit", L"Skala an Daten anpassen");
    xml += std::format(L"<menu id=\"rtkScales_{}\" label=\"Skalen\">", suffix) + toggle(L"rtkT_scaleYear", L"Jahr") +
           toggle(L"rtkT_scaleQuarter", L"Quartal") + toggle(L"rtkT_scaleMonth", L"Monat") + toggle(L"rtkT_scaleWeek", L"Woche") +
           toggle(L"rtkT_scaleDay", L"Tag") + button(L"rtkG_autoScales", L"Automatische Skalen") + L"</menu>";
    xml += std::format(L"<menu id=\"rtkView_{}\" label=\"Anzeige\">", suffix) + toggle(L"rtkT_weekendsShaded", L"Wochenenden schattieren") +
           toggle(L"rtkT_weekendsHidden", L"Wochenenden ausblenden") + toggle(L"rtkT_separators", L"Trennlinien") +
           toggle(L"rtkT_barLabels", L"Datum an Balken") + toggle(L"rtkT_milestoneLabels", L"Datum an Meilensteinen") +
           toggle(L"rtkT_milestoneLines", L"Meilensteinlinien") + toggle(L"rtkT_today", L"Heute-Linie") +
           toggle(L"rtkT_rowSeparators", L"Zeilentrennlinien") + toggle(L"rtkT_headlines", L"Spaltenüberschriften") +
           toggle(L"rtkT_responsible", L"Spalte „Verantwortlich“") + toggle(L"rtkT_remark", L"Spalte „Bemerkung“") + L"</menu>";
    xml += separator(3);
    xml += button(L"rtkG_datasheet", L"Datenblatt öffnen");
    return xml + L"</menu>";
}

// "rtkG_insertRow_a" → "rtkG_insertRow"
std::wstring baseId(const std::wstring& id) {
    const size_t first = id.find(L'_');
    const size_t second = first == std::wstring::npos ? std::wstring::npos : id.find(L'_', first + 1);
    return second == std::wstring::npos ? id : id.substr(0, second);
}

GanttHit g_contextHit;  // Mausposition beim Öffnen des Kontextmenüs

long longArg(const DISPPARAMS* params, UINT index) {
    if (!params || index >= params->cArgs) throw Error("Argument fehlt");
    VARIANT value;
    VariantInit(&value);
    check(VariantCopyInd(&value, &params->rgvarg[params->cArgs - 1 - index]), "Argument lesen");
    VARIANT number;
    VariantInit(&number);
    const HRESULT hr = VariantChangeType(&number, &value, 0, VT_I4);
    VariantClear(&value);
    check(hr, "Argument als Zahl lesen");
    return number.lVal;
}

// Optionales Wahrheitswert-Argument; fehlt es, gilt `fallback`.
bool boolArg(const DISPPARAMS* params, UINT index, bool fallback) {
    if (!params || index >= params->cArgs) return fallback;
    VARIANT value;
    VariantInit(&value);
    if (FAILED(VariantCopyInd(&value, &params->rgvarg[params->cArgs - 1 - index]))) return fallback;
    if (value.vt == VT_ERROR || value.vt == VT_EMPTY) {
        VariantClear(&value);
        return fallback;
    }
    VARIANT flag;
    VariantInit(&flag);
    const HRESULT hr = VariantChangeType(&flag, &value, 0, VT_BOOL);
    VariantClear(&value);
    return SUCCEEDED(hr) ? flag.boolVal != VARIANT_FALSE : fallback;
}

std::wstring loadRibbonXml() {
    const auto module = reinterpret_cast<HMODULE>(&__ImageBase);
    HRSRC info = FindResourceW(module, MAKEINTRESOURCEW(IDR_RIBBON), MAKEINTRESOURCEW(10) /* RT_RCDATA */);
    if (!info) return {};
    HGLOBAL data = LoadResource(module, info);
    const auto* bytes = static_cast<const char*>(LockResource(data));
    std::string_view xml(bytes, SizeofResource(module, info));
    if (xml.starts_with("\xEF\xBB\xBF")) xml.remove_prefix(3);
    std::wstring text = widen(xml);
    for (const wchar_t* suffix : {L"a", L"b"}) {
        const std::wstring marker = std::format(L"<!--rethink:gantt-menu {}-->", suffix);
        if (const size_t at = text.find(marker); at != std::wstring::npos) text.replace(at, marker.size(), ganttMenuXml(suffix));
    }
    return text;
}

// Steuerelement eines Ribbon-Rückrufs (Argument `index`) → seine ID
std::wstring controlId(const DISPPARAMS* params, UINT index) {
    if (!params || index >= params->cArgs) return {};
    const VARIANT& v = params->rgvarg[params->cArgs - 1 - index];
    IDispatch* control = v.vt == VT_DISPATCH ? v.pdispVal : v.vt == (VT_DISPATCH | VT_BYREF) && v.ppdispVal ? *v.ppdispVal : nullptr;
    if (!control) return {};
    try {
        return disp::call(control, L"Id").str();
    } catch (...) {
        return {};
    }
}

// Rückgabe eines get…-Rückrufs: über pVarResult und, falls Office ein Ausgabeargument mitgibt, auch dorthin.
void setBool(VARIANT* result, DISPPARAMS* params, bool value) {
    if (result) {
        VariantInit(result);
        result->vt = VT_BOOL;
        result->boolVal = value ? VARIANT_TRUE : VARIANT_FALSE;
    }
    if (params && params->cArgs >= 2 && params->rgvarg[0].vt == (VT_BOOL | VT_BYREF) && params->rgvarg[0].pboolVal)
        *params->rgvarg[0].pboolVal = value ? VARIANT_TRUE : VARIANT_FALSE;
}

engine::Orientation orientationArg(const std::wstring& text) {
    if (text == L"down") return engine::Orientation::Down;
    if (text == L"right") return engine::Orientation::Right;
    if (text == L"left") return engine::Orientation::Left;
    return engine::Orientation::Up;
}

ElementKind kindArg(const std::wstring& text) {
    if (text == L"gantt") return ElementKind::Gantt;
    if (text == L"waterfall-down") return ElementKind::WaterfallDown;
    if (text == L"waterfall") return ElementKind::Waterfall;
    throw Error("Elementart: \"waterfall\", \"waterfall-down\" oder \"gantt\"");
}

double doubleArg(const DISPPARAMS* params, UINT index) {
    if (!params || index >= params->cArgs) throw Error("Argument fehlt");
    VARIANT value;
    VariantInit(&value);
    check(VariantCopyInd(&value, &params->rgvarg[params->cArgs - 1 - index]), "Argument lesen");
    VARIANT number;
    VariantInit(&number);
    const HRESULT hr = VariantChangeType(&number, &value, 0, VT_R8);
    VariantClear(&value);
    check(hr, "Argument als Zahl lesen");
    return number.dblVal;
}

// Optionale Folienkoordinaten (x, y) ab Argument `index`: Treffer im Zeitplan; ohne Koordinaten ein leerer Treffer.
GanttHit hitArg(IDispatch* app, const DISPPARAMS* params, UINT index) {
    if (!params || index + 1 >= params->cArgs) return {};
    return ganttHitAtPoint(app, doubleArg(params, index), doubleArg(params, index + 1));
}

std::wstring hitText(const GanttHit& h) {
    return std::format(L"gueltig={} zeile={} datum={} element={} skala={}", h.valid, h.row, h.hasDate ? engine::isoDate(h.date) : L"-",
                       h.role.empty() ? L"-" : widen(h.role), h.inScale);
}

template <class F>
void ribbonGuarded(F&& body) {
    try {
        body();
    } catch (const winrt::hresult_error& e) {
        MessageBoxW(nullptr, e.message().c_str(), L"rethink", MB_ICONINFORMATION);
    } catch (const std::exception& e) {
        MessageBoxW(nullptr, widen(e.what()).c_str(), L"rethink", MB_ICONINFORMATION);
    } catch (...) {
    }
}

// Argument `index` (0 = erstes) als Zeichenkette; fehlende optionale Argumente ergeben "".
std::wstring stringArg(const DISPPARAMS* params, UINT index, bool optional) {
    if (!params || index >= params->cArgs) {
        if (optional) return {};
        throw Error("Argument fehlt");
    }
    const VARIANT& source = params->rgvarg[params->cArgs - 1 - index];
    VARIANT value;
    VariantInit(&value);
    check(VariantCopyInd(&value, &source), "Argument lesen");
    if (value.vt == VT_ERROR || value.vt == VT_EMPTY) {
        VariantClear(&value);
        if (optional) return {};
        throw Error("Argument fehlt");
    }
    VARIANT text;
    VariantInit(&text);
    const HRESULT hr = VariantChangeType(&text, &value, 0, VT_BSTR);
    VariantClear(&value);
    check(hr, "Argument als Text lesen");
    std::wstring s(text.bstrVal, SysStringLen(text.bstrVal));
    VariantClear(&text);
    return s;
}

// Mausweg-Liste "10;25;-5" (Pixel, positiv = nach oben) für DragSegment.
std::vector<int> pixelListArg(const DISPPARAMS* params, UINT index) {
    std::vector<int> pixels;
    const std::wstring text = stringArg(params, index, false);
    for (size_t pos = 0; pos <= text.size();) {
        const size_t end = std::min(text.find(L';', pos), text.size());
        const std::wstring part = text.substr(pos, end - pos);
        if (!part.empty()) {
            try {
                pixels.push_back(std::stoi(part));
            } catch (const std::exception&) {
                throw Error("Mausweg als Zahlen mit ; getrennt angeben, z. B. \"10;25;-5\"");
            }
        }
        pos = end + 1;
    }
    return pixels;
}

void setString(VARIANT* result, std::wstring_view s) {
    if (!result) return;
    VariantInit(result);
    result->vt = VT_BSTR;
    result->bstrVal = SysAllocStringLen(s.data(), static_cast<UINT>(s.size()));
}

void setLong(VARIANT* result, long v) {
    if (!result) return;
    VariantInit(result);
    result->vt = VT_I4;
    result->lVal = v;
}

HRESULT reportException(EXCEPINFO* info, const std::wstring& description) {
    if (info) {
        *info = {};
        info->scode = E_FAIL;
        info->bstrSource = SysAllocString(L"rethink");
        info->bstrDescription = SysAllocString(description.c_str());
    }
    return DISP_E_EXCEPTION;
}

class __declspec(uuid("68EFFB93-253D-4A83-AA32-87F9C46D9023")) Addin
    : public RuntimeClass<RuntimeClassFlags<ClassicCom>, ChainInterfaces<IDTExtensibility2, IDispatch>, IRibbonExtensibility> {
public:
    // --- IDispatch (gemeinsam für beide Schnittstellen) ---

    STDMETHODIMP GetTypeInfoCount(UINT* count) override {
        if (!count) return E_POINTER;
        *count = 0;
        return S_OK;
    }

    STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo** info) override {
        if (info) *info = nullptr;
        return E_NOTIMPL;
    }

    STDMETHODIMP GetIDsOfNames(REFIID, LPOLESTR* names, UINT count, LCID, DISPID* ids) override {
        if (!names || !ids) return E_POINTER;
        HRESULT hr = S_OK;
        for (UINT i = 0; i < count; ++i) {
            ids[i] = DISPID_UNKNOWN;
            if (i == 0) {
                for (auto& m : kMembers)
                    if (_wcsicmp(names[0], m.name) == 0) ids[0] = m.id;
            }
            if (ids[i] == DISPID_UNKNOWN) hr = DISP_E_UNKNOWNNAME;  // benannte Argumente werden nicht unterstützt
        }
        return hr;
    }

    STDMETHODIMP Invoke(DISPID id, REFIID, LCID, WORD, DISPPARAMS* params, VARIANT* result, EXCEPINFO* excepInfo, UINT*) override {
        switch (id) {
        case kOnBuild: commandBuild(app_.Get()); return S_OK;
        case kOnInsertChart: commandInsertChart(app_.Get()); return S_OK;
        case kOnSetName: commandSetName(app_.Get()); return S_OK;
        case kOnInsertWaterfall: commandInsertWaterfall(app_.Get()); return S_OK;
        case kOnOpenDatasheet: commandOpenDatasheet(app_.Get()); return S_OK;
        case kOnInsertGantt: commandInsertElement(app_.Get(), ElementKind::Gantt, engine::Orientation::Up); return S_OK;
        case kOnInsertGallery: {
            POINT at{};
            GetCursorPos(&at);
            ComPtr<IDispatch> app = app_;
            HWND owner = nullptr;
            try {
                owner = reinterpret_cast<HWND>(static_cast<LONG_PTR>(disp::call(app.Get(), L"HWND").toLong()));
            } catch (...) {
                owner = GetActiveWindow();
            }
            openGallery(owner, at, [app](GalleryChoice choice) { commandInsertElement(app.Get(), choice.kind, choice.orientation); });
            return S_OK;
        }
        case kOnGanttVisible: {
            bool visible = false;
            try {
                g_contextHit = app_ ? ganttHitAtCursor(app_.Get()) : GanttHit{};
                visible = g_contextHit.valid;
            } catch (...) {
            }
            setBool(result, params, visible);
            return S_OK;
        }
        case kOnGanttAction: {
            const std::wstring control = baseId(controlId(params, 0));
            ribbonGuarded([&] { ganttCommand(app_.Get(), control, g_contextHit); });
            return S_OK;
        }
        case kOnGanttToggle: {
            const std::wstring control = baseId(controlId(params, 0));
            const bool pressed = boolArg(params, 1, false);
            ribbonGuarded([&] { ganttToggle(app_.Get(), control, pressed, g_contextHit); });
            return S_OK;
        }
        case kOnGanttPressed: {
            bool pressed = false;
            try {
                pressed = ganttPressed(app_.Get(), baseId(controlId(params, 0)), g_contextHit);
            } catch (...) {
            }
            setBool(result, params, pressed);
            return S_OK;
        }
        case kOnRibbonLoad:
            if (params && params->cArgs >= 1 && params->rgvarg[0].vt == VT_DISPATCH) ribbon_ = params->rgvarg[0].pdispVal;
            return S_OK;
        case kRibbonStatus: setString(result, ribbon_ ? L"geladen" : L"nicht geladen"); return S_OK;
        default: break;
        }

        try {
            if (id == kVersion) {
                setString(result, L"" RETHINK_VERSION);
                return S_OK;
            }
            if (id == kEventStatus) {
                setString(result, eventStatus());
                return S_OK;
            }
            if (id == kLastWarnings) {
                setString(result, lastWarnings_);
                return S_OK;
            }
            if (!app_) throw Error("Das Add-in ist nicht mit PowerPoint verbunden");
            switch (id) {
            case kBuildFrom: {
                std::vector<std::string> warnings;
                const auto opened = buildAndOpen(app_.Get(), stringArg(params, 0, false), stringArg(params, 1, true), warnings);
                lastWarnings_.clear();
                for (auto& w : warnings) lastWarnings_ += (lastWarnings_.empty() ? L"" : L"\n") + widen(w);
                setString(result, opened);
                return S_OK;
            }
            case kInsertChart: setString(result, insertChart(app_.Get())); return S_OK;
            case kSetName: setLong(result, setElementName(app_.Get(), stringArg(params, 0, false))); return S_OK;
            case kGetName: setString(result, elementName(app_.Get())); return S_OK;
            case kInsertWaterfall: setString(result, insertWaterfall(app_.Get(), boolArg(params, 0, false))); return S_OK;
            case kOpenDatasheet: openDatasheet(app_.Get(), boolArg(params, 0, false)); return S_OK;
            case kOpenExcelDatasheet: setString(result, openExcelDatasheet(app_.Get(), boolArg(params, 0, false))); return S_OK;
            case kDatasheetEnter: datasheetEnterCell(app_.Get(), longArg(params, 0), longArg(params, 1), stringArg(params, 2, true)); return S_OK;
            case kDatasheetPaste: datasheetPasteText(app_.Get(), longArg(params, 0), longArg(params, 1), stringArg(params, 2, false)); return S_OK;
            case kDatasheetCell: setString(result, datasheetCellText(app_.Get(), longArg(params, 0), longArg(params, 1))); return S_OK;
            case kCloseDatasheet: closeDatasheet(app_.Get()); return S_OK;
            case kDatasheetCount: setLong(result, openDatasheetCount()); return S_OK;
            case kApplyDatasheets: setLong(result, applyDatasheets(app_.Get())); return S_OK;
            case kAdjustSegment:
                adjustSegment(app_.Get(), longArg(params, 0), longArg(params, 1), longArg(params, 2) >= 0 ? 1 : -1, boolArg(params, 3, false));
                return S_OK;
            case kOverlayInfo: setString(result, overlayInfo()); return S_OK;
            case kDragSegment: setString(result, simulateSegmentDrag(longArg(params, 0), longArg(params, 1), pixelListArg(params, 2))); return S_OK;
            case kInsertElement:
                setString(result, insertElement(app_.Get(), kindArg(stringArg(params, 0, false)), orientationArg(stringArg(params, 1, true)),
                                                boolArg(params, 2, false)));
                return S_OK;
            case kGanttTargets: setString(result, ganttTargets(app_.Get())); return S_OK;
            case kGanttDrag:
                setString(result, simulateGanttDrag(narrow(stringArg(params, 0, false)), longArg(params, 1), pixelListArg(params, 2),
                                                    boolArg(params, 3, false)));
                return S_OK;
            case kGanttCommand: ganttCommand(app_.Get(), stringArg(params, 0, false), hitArg(app_.Get(), params, 1)); return S_OK;
            case kGanttToggle: ganttToggle(app_.Get(), stringArg(params, 0, false), boolArg(params, 1, false), hitArg(app_.Get(), params, 2)); return S_OK;
            case kGanttPressed: setBool(result, nullptr, ganttPressed(app_.Get(), stringArg(params, 0, false), hitArg(app_.Get(), params, 1))); return S_OK;
            case kSetGanttDates:
                setGanttDates(app_.Get(), narrow(stringArg(params, 0, true)), stringArg(params, 1, false), stringArg(params, 2, true));
                return S_OK;
            case kGanttMenu: setString(result, ganttMenuText(app_.Get(), hitArg(app_.Get(), params, 0))); return S_OK;
            case kGanttHit: setString(result, hitText(ganttHitAtPoint(app_.Get(), doubleArg(params, 0), doubleArg(params, 1)))); return S_OK;
            case kOpenCalendar: openGanttCalendar(app_.Get(), narrow(stringArg(params, 0, true))); return S_OK;
            case kCloseCalendar: setBool(result, nullptr, calendarWindowOpen()); closeCalendarWindow(); return S_OK;
            case kActivateTab:
                if (!ribbon_) throw Error("Das Ribbon wurde nicht geladen (Ribbon-XML abgelehnt?)");
                disp::call(ribbon_.Get(), L"ActivateTab", std::wstring(L"rethinkTab"));
                return S_OK;
            case kDatasheetWorkbook: {
                auto workbook = datasheetWorkbook(app_.Get());
                if (result) {
                    VariantInit(result);
                    result->vt = VT_DISPATCH;
                    result->pdispVal = workbook.Detach();
                }
                return S_OK;
            }
            default: return DISP_E_MEMBERNOTFOUND;
            }
        } catch (const winrt::hresult_error& e) {
            return reportException(excepInfo, e.message().c_str());
        } catch (const std::exception& e) {
            return reportException(excepInfo, widen(e.what()));
        } catch (...) {
            return reportException(excepInfo, L"Unerwarteter Fehler");
        }
    }

    // --- IDTExtensibility2 ---

    STDMETHODIMP OnConnection(IDispatch* application, int, IDispatch* addInInst, SAFEARRAY**) override {
        app_ = application;
        try {
            // Macht dieses Objekt als COMAddIns("rethink.addin").Object für Automation erreichbar.
            if (addInInst) disp::put(addInInst, L"Object", static_cast<IDispatch*>(static_cast<IDTExtensibility2*>(this)));
        } catch (...) {
        }
        try {
            startEvents(application);  // Leuchtrahmen bei Auswahl, Doppelklick öffnet das Datenblatt
        } catch (...) {
        }
        return S_OK;
    }

    STDMETHODIMP OnDisconnection(int, SAFEARRAY**) override {
        shutdown();
        app_.Reset();
        return S_OK;
    }

    STDMETHODIMP OnAddInsUpdate(SAFEARRAY**) override { return S_OK; }
    STDMETHODIMP OnStartupComplete(SAFEARRAY**) override { return S_OK; }
    STDMETHODIMP OnBeginShutdown(SAFEARRAY**) override {
        shutdown();
        return S_OK;
    }

    static void shutdown() {
        try {
            stopEvents();
            stopDatasheetWatch();
        } catch (...) {
        }
    }

    // --- IRibbonExtensibility ---

    STDMETHODIMP GetCustomUI(BSTR ribbonId, BSTR* ribbonXml) override {
        if (!ribbonXml) return E_POINTER;
        *ribbonXml = nullptr;
        if (!ribbonId || _wcsicmp(ribbonId, L"Microsoft.PowerPoint.Presentation") != 0) return S_OK;
        try {
            const auto xml = loadRibbonXml();
            *ribbonXml = SysAllocStringLen(xml.data(), static_cast<UINT>(xml.size()));
            return *ribbonXml ? S_OK : E_OUTOFMEMORY;
        } catch (...) {
            return E_FAIL;
        }
    }

private:
    ComPtr<IDispatch> app_;
    ComPtr<IDispatch> ribbon_;  // IRibbonUI aus onLoad
    std::wstring lastWarnings_;
};

CoCreatableClass(Addin);

}  // namespace
}  // namespace rethink::addin

STDAPI DllGetClassObject(REFCLSID clsid, REFIID iid, void** object) {
    return Microsoft::WRL::Module<Microsoft::WRL::InProc>::GetModule().GetClassObject(clsid, iid, object);
}

STDAPI DllCanUnloadNow() {
    return Microsoft::WRL::Module<Microsoft::WRL::InProc>::GetModule().GetObjectCount() == 0 ? S_OK : S_FALSE;
}
