// COM-Add-in für PowerPoint: IDTExtensibility2 (Laden/Entladen) und IRibbonExtensibility (Ribbon-XML).
// Ribbon-Rückrufe und die Automationsschnittstelle laufen über dieselbe IDispatch-Implementierung.

#include <winrt/base.h>

#include "commands.h"
#include "dispatch.h"
#include "elements.h"
#include "events.h"
#include "officeapi.h"
#include "resource.h"
#include "version.h"

#include <wrl/implements.h>
#include <wrl/module.h>

// WRL (module.h, StaticStorage) löst bei der Instanziierung C4324 aus: Auffüllung wegen alignas – unkritisch.
#pragma warning(disable : 4324)

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
    {L"DatasheetCount", kDatasheetCount},
};

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
    return widen(xml);
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
