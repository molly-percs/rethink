#include "events.h"

#include "dispatch.h"
#include "elements.h"
#include "overlay.h"

#include <ocidl.h>

#include <format>
#include <string>

namespace rethink::addin {
namespace {

// PowerPoint EApplication (Ereignisschnittstelle des Application-Objekts)
constexpr GUID kDiidEApplication = {0x914934C2, 0x5A91, 0x11CF, {0x87, 0x00, 0x00, 0xAA, 0x00, 0x60, 0x26, 0x3B}};

ComPtr<IDispatch> g_app;
ComPtr<IConnectionPoint> g_point;
DWORD g_cookie = 0;
GUID g_eventIid = kDiidEApplication;  // tatsächlich verwendete Ereignisschnittstelle
long g_eventCount = 0;
std::wstring g_eventError;
DISPID g_selectionChange = 2001, g_beforeDoubleClick = 2003;  // überschrieben aus der Typbibliothek

ComPtr<IDispatch> g_element;  // aktuell markiertes rethink-Element
RECT g_lastRect{};
UINT_PTR g_trackTimer = 0;

void stopTracking() {
    g_element.Reset();
    g_lastRect = {};
    hideOverlay();
    if (g_trackTimer) KillTimer(nullptr, g_trackTimer);
    g_trackTimer = 0;
}

// Hauptfenster von PowerPoint: Application.HWND ist nicht in jeder Version per IDispatch erreichbar,
// daher das sichtbare Rahmenfenster (PPTFrameClass) des eigenen UI-Threads.
HWND powerPointWindow() {
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

// Lage des Elements in Bildschirmpixeln (folgt Zoom und Bildlauf).
void placeOverlay() {
    auto window = disp::object(g_app.Get(), L"ActiveWindow");
    const double left = disp::call(g_element.Get(), L"Left").toDouble(), top = disp::call(g_element.Get(), L"Top").toDouble();
    const double right = left + disp::call(g_element.Get(), L"Width").toDouble();
    const double bottom = top + disp::call(g_element.Get(), L"Height").toDouble();
    RECT r{disp::call(window.Get(), L"PointsToScreenPixelsX", left).toLong(), disp::call(window.Get(), L"PointsToScreenPixelsY", top).toLong(),
           disp::call(window.Get(), L"PointsToScreenPixelsX", right).toLong(), disp::call(window.Get(), L"PointsToScreenPixelsY", bottom).toLong()};
    if (EqualRect(&r, &g_lastRect) && overlayVisible()) return;
    g_lastRect = r;
    showOverlay(powerPointWindow(), r, elementLabel(g_element.Get()));
}

void remember(const char* where, const std::exception& e) {
    g_eventError = widen(where) + L": " + widen(e.what());
}

void CALLBACK onTrack(HWND, UINT, UINT_PTR, DWORD) {
    if (!g_element) return stopTracking();
    try {
        // Element noch markiert? (z. B. nach Neuzeichnen aus dem Datenblatt ist es eine neue Gruppe)
        auto current = rethinkElementOf(disp::object(disp::object(g_app.Get(), L"ActiveWindow").Get(), L"Selection").Get());
        if (!current) return stopTracking();
        g_element = current;
        placeOverlay();
    } catch (const std::exception& e) {
        remember("Nachführen", e);
        stopTracking();
    } catch (...) {
        stopTracking();
    }
}

void onSelectionChange(IDispatch* selection) {
    try {
        auto element = rethinkElementOf(selection);
        if (!element) return stopTracking();
        g_element = element;
        g_lastRect = {};
        placeOverlay();
        if (!g_trackTimer) g_trackTimer = SetTimer(nullptr, 0, 150, onTrack);
    } catch (const std::exception& e) {
        remember("Auswahl", e);
        stopTracking();
    } catch (...) {
        stopTracking();
    }
}

// true, wenn der Doppelklick einem rethink-Element galt und übernommen wurde.
bool onBeforeDoubleClick(IDispatch* selection) {
    if (!rethinkElementOf(selection)) return false;
    commandOpenDatasheet(g_app.Get());
    return true;
}

IDispatch* dispatchArg(const VARIANT& v) {
    if (v.vt == VT_DISPATCH) return v.pdispVal;
    if (v.vt == (VT_DISPATCH | VT_BYREF) && v.ppdispVal) return *v.ppdispVal;
    return nullptr;
}

class EventSink final : public IDispatch {
public:
    STDMETHODIMP QueryInterface(REFIID iid, void** out) override {
        if (!out) return E_POINTER;
        if (iid == IID_IUnknown || iid == IID_IDispatch || iid == g_eventIid) {
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
    STDMETHODIMP GetTypeInfoCount(UINT* count) override {
        if (count) *count = 0;
        return S_OK;
    }
    STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo** info) override {
        if (info) *info = nullptr;
        return E_NOTIMPL;
    }
    STDMETHODIMP GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*) override { return E_NOTIMPL; }
    STDMETHODIMP Invoke(DISPID id, REFIID, LCID, WORD, DISPPARAMS* params, VARIANT*, EXCEPINFO*, UINT*) override {
        ++g_eventCount;
        try {
            if (id == g_selectionChange && params && params->cArgs >= 1) {
                onSelectionChange(dispatchArg(params->rgvarg[params->cArgs - 1]));
            } else if (id == g_beforeDoubleClick && params && params->cArgs >= 2) {
                // Argumente rückwärts: rgvarg[1] = Sel, rgvarg[0] = Cancel (VARIANT_BOOL*)
                if (onBeforeDoubleClick(dispatchArg(params->rgvarg[1]))) {
                    VARIANT& cancel = params->rgvarg[0];
                    if (cancel.vt == (VT_BOOL | VT_BYREF) && cancel.pboolVal) *cancel.pboolVal = VARIANT_TRUE;
                }
            }
        } catch (...) {
            // Ereignisse dürfen PowerPoint nie mit Ausnahmen erreichen.
        }
        return S_OK;
    }

private:
    LONG refs_ = 1;
};

// Ereignis-IDs aus der Typbibliothek von PowerPoint lesen (robust gegen Versionsunterschiede).
void resolveDispIds(IDispatch* app) {
    ComPtr<ITypeInfo> appInfo;
    if (FAILED(app->GetTypeInfo(0, LOCALE_USER_DEFAULT, &appInfo)) || !appInfo) return;
    ComPtr<ITypeLib> lib;
    UINT index = 0;
    if (FAILED(appInfo->GetContainingTypeLib(&lib, &index))) return;
    ComPtr<ITypeInfo> events;
    if (FAILED(lib->GetTypeInfoOfGuid(g_eventIid, &events))) return;
    LPOLESTR names[] = {const_cast<LPOLESTR>(L"WindowSelectionChange"), const_cast<LPOLESTR>(L"WindowBeforeDoubleClick")};
    DISPID id = 0;
    if (SUCCEEDED(events->GetIDsOfNames(&names[0], 1, &id))) g_selectionChange = id;
    if (SUCCEEDED(events->GetIDsOfNames(&names[1], 1, &id))) g_beforeDoubleClick = id;
}

}  // namespace

void startEvents(IDispatch* app) {
    stopEvents();
    if (!app) return;
    g_app = app;
    g_eventError.clear();
    ComPtr<IConnectionPointContainer> container;
    HRESULT hr = app->QueryInterface(IID_PPV_ARGS(&container));
    if (FAILED(hr)) {
        g_eventError = std::format(L"kein IConnectionPointContainer (0x{:08X})", static_cast<unsigned>(hr));
        return;
    }
    // Bekannte Schnittstelle versuchen; sonst den (einzigen) Verbindungspunkt des Application-Objekts nehmen.
    if (FAILED(container->FindConnectionPoint(kDiidEApplication, &g_point))) {
        ComPtr<IEnumConnectionPoints> points;
        if (SUCCEEDED(container->EnumConnectionPoints(&points))) {
            ComPtr<IConnectionPoint> point;
            ULONG fetched = 0;
            if (points->Next(1, &point, &fetched) == S_OK) g_point = point;
        }
    }
    if (!g_point) {
        g_eventError = L"kein Verbindungspunkt für Ereignisse";
        return;
    }
    g_point->GetConnectionInterface(&g_eventIid);
    resolveDispIds(app);
    auto* sink = new EventSink();
    hr = g_point->Advise(sink, &g_cookie);
    if (FAILED(hr)) {
        g_eventError = std::format(L"Advise fehlgeschlagen (0x{:08X})", static_cast<unsigned>(hr));
        g_point.Reset();
        g_cookie = 0;
    }
    sink->Release();  // die Verbindung hält ihre eigene Referenz
}

std::wstring eventStatus() {
    wchar_t iid[64] = {};
    StringFromGUID2(g_eventIid, iid, 64);
    return std::format(L"verbunden={} iid={} auswahl={} doppelklick={} ereignisse={} overlay={} datenblaetter={}{}", g_cookie != 0, iid,
                       g_selectionChange, g_beforeDoubleClick, g_eventCount, overlayVisible(), watchCount(),
                       g_eventError.empty() ? L"" : L" fehler=" + g_eventError);
}

void stopEvents() {
    stopTracking();
    if (g_point && g_cookie) g_point->Unadvise(g_cookie);
    g_point.Reset();
    g_cookie = 0;
    g_app.Reset();
}

}  // namespace rethink::addin
