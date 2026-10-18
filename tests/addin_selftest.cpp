// Selbsttest des Add-ins ohne Registrierung und ohne PowerPoint:
// lädt die DLL direkt, erzeugt das COM-Objekt über DllGetClassObject und prüft die Schnittstellen.
//   addin_selftest <pfad\rethink-addin.dll>

#include "officeapi.h"

#include <wrl/client.h>

#include <algorithm>
#include <cstdio>
#include <string>
#include <vector>

using Microsoft::WRL::ComPtr;

namespace {

constexpr CLSID kClsid = {0x68EFFB93, 0x253D, 0x4A83, {0xAA, 0x32, 0x87, 0xF9, 0xC4, 0x6D, 0x90, 0x23}};
int failures = 0;

void expect(bool ok, const char* what) {
    std::printf("%s %s\n", ok ? "ok  " : "FEHLER", what);
    if (!ok) ++failures;
}

DISPID idOf(IDispatch* d, const wchar_t* name) {
    LPOLESTR names[] = {const_cast<LPOLESTR>(name)};
    DISPID id = DISPID_UNKNOWN;
    return SUCCEEDED(d->GetIDsOfNames(IID_NULL, names, 1, LOCALE_USER_DEFAULT, &id)) ? id : DISPID_UNKNOWN;
}

}  // namespace

int wmain(int argc, wchar_t** argv) {
    if (argc != 2) {
        std::fputs("Aufruf: addin_selftest <rethink-addin.dll>\n", stderr);
        return 1;
    }
    SetConsoleOutputCP(CP_UTF8);
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    HMODULE dll = LoadLibraryW(argv[1]);
    expect(dll != nullptr, "DLL laden");
    if (!dll) return 1;

    using GetClassObjectFn = HRESULT(STDAPICALLTYPE*)(REFCLSID, REFIID, void**);
    using CanUnloadFn = HRESULT(STDAPICALLTYPE*)();
    auto getClassObject = reinterpret_cast<GetClassObjectFn>(GetProcAddress(dll, "DllGetClassObject"));
    auto canUnload = reinterpret_cast<CanUnloadFn>(GetProcAddress(dll, "DllCanUnloadNow"));
    expect(getClassObject && canUnload, "Exporte DllGetClassObject/DllCanUnloadNow");
    if (!getClassObject || !canUnload) return 1;

    {
        ComPtr<IClassFactory> factory;
        expect(SUCCEEDED(getClassObject(kClsid, IID_PPV_ARGS(&factory))), "Klassenfabrik für die CLSID");
        CLSID other = kClsid;
        other.Data1 ^= 1;
        ComPtr<IClassFactory> none;
        expect(FAILED(getClassObject(other, IID_PPV_ARGS(&none))), "fremde CLSID wird abgelehnt");

        ComPtr<IDTExtensibility2> ext;
        expect(factory && SUCCEEDED(factory->CreateInstance(nullptr, IID_PPV_ARGS(&ext))), "Objekt erzeugen (IDTExtensibility2)");
        if (!ext) return 1;
        ComPtr<IRibbonExtensibility> ribbon;
        ComPtr<IDispatch> dispatch;
        expect(SUCCEEDED(ext.As(&ribbon)), "QueryInterface IRibbonExtensibility");
        expect(SUCCEEDED(ext.As(&dispatch)), "QueryInterface IDispatch");
        expect(canUnload() == S_FALSE, "DllCanUnloadNow: S_FALSE solange Objekte leben");

        BSTR xml = nullptr;
        BSTR id = SysAllocString(L"Microsoft.PowerPoint.Presentation");
        expect(SUCCEEDED(ribbon->GetCustomUI(id, &xml)) && xml, "GetCustomUI liefert XML");
        const std::wstring text = xml ? xml : L"";
        expect(text.find(L"label=\"rethink\"") != std::wstring::npos, "Ribbon-XML: Tab \"rethink\"");
        expect(text.find(L"Aus .ppttc erzeugen…") != std::wstring::npos, "Ribbon-XML: Umlaute/Auslassung korrekt (UTF-8)");
        for (const wchar_t* cb : {L"RethinkBuild", L"RethinkInsertChart", L"RethinkSetName", L"RethinkInsertWaterfall", L"RethinkOpenDatasheet",
                                  L"RethinkInsertGallery", L"RethinkInsertGantt", L"RethinkGanttAction", L"RethinkGanttToggle"})
            expect(text.find(std::wstring(L"onAction=\"") + cb + L"\"") != std::wstring::npos && idOf(dispatch.Get(), cb) != DISPID_UNKNOWN,
                   "Rückruf im XML und per GetIDsOfNames auflösbar");
        for (const wchar_t* cb : {L"RethinkGanttVisible", L"RethinkGanttPressed"})
            expect(text.find(cb) != std::wstring::npos && idOf(dispatch.Get(), cb) != DISPID_UNKNOWN, "get-Rückruf im XML und auflösbar");
        expect(text.find(L"<!--rethink:gantt-menu") == std::wstring::npos && text.find(L"id=\"rtkGantt_a\"") != std::wstring::npos &&
                   text.find(L"id=\"rtkGantt_b\"") != std::wstring::npos && text.find(L"id=\"rtkG_insertRow_b\"") != std::wstring::npos,
               "Kontextmenü „rethink-Zeitplan“ in beiden PowerPoint-Kontextmenüs, IDs eindeutig");
        {
            // Steuerelement-IDs müssen eindeutig sein
            std::vector<std::wstring> ids;
            for (size_t at = text.find(L" id=\""); at != std::wstring::npos; at = text.find(L" id=\"", at + 1)) {
                const size_t end = text.find(L'"', at + 5);
                ids.push_back(text.substr(at + 5, end - at - 5));
            }
            std::sort(ids.begin(), ids.end());
            expect(std::adjacent_find(ids.begin(), ids.end()) == ids.end() && ids.size() > 60, "alle Steuerelement-IDs eindeutig");
        }
        SysFreeString(xml);
        SysFreeString(id);

        expect(idOf(dispatch.Get(), L"insertchart") != DISPID_UNKNOWN, "GetIDsOfNames ohne Groß-/Kleinschreibung");
        expect(idOf(dispatch.Get(), L"GibtEsNicht") == DISPID_UNKNOWN, "unbekannter Name wird abgelehnt");

        // Version (ohne Verbindung zu PowerPoint)
        DISPPARAMS none0{};
        VARIANT result;
        VariantInit(&result);
        expect(SUCCEEDED(dispatch->Invoke(idOf(dispatch.Get(), L"Version"), IID_NULL, 0, DISPATCH_METHOD, &none0, &result, nullptr, nullptr)) &&
                   result.vt == VT_BSTR && SysStringLen(result.bstrVal) > 0,
               "Version");
        VariantClear(&result);

        // Ohne OnConnection muss ein Befehl sauber als Ausnahme gemeldet werden, nicht abstürzen.
        EXCEPINFO info{};
        const HRESULT hr = dispatch->Invoke(idOf(dispatch.Get(), L"InsertChart"), IID_NULL, 0, DISPATCH_METHOD, &none0, &result, &info, nullptr);
        expect(hr == DISP_E_EXCEPTION && info.bstrDescription, "Befehl ohne PowerPoint: DISP_E_EXCEPTION mit Text");
        SysFreeString(info.bstrSource);
        SysFreeString(info.bstrDescription);
        SysFreeString(info.bstrHelpFile);

        expect(SUCCEEDED(ext->OnConnection(nullptr, 0, nullptr, nullptr)) && SUCCEEDED(ext->OnDisconnection(0, nullptr)),
               "OnConnection/OnDisconnection ohne Argumente");
    }
    expect(canUnload() == S_OK, "DllCanUnloadNow: S_OK nach Freigabe");
    FreeLibrary(dll);
    std::printf("%s (%d Fehler)\n", failures ? "FEHLGESCHLAGEN" : "BESTANDEN", failures);
    return failures ? 2 : 0;
}
