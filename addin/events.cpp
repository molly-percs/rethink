#include <winrt/base.h>

#include "events.h"

#include "dispatch.h"
#include "elements.h"
#include "engine/chart.h"
#include "overlay.h"

#include <ocidl.h>

#include <cmath>
#include <format>
#include <string>
#include <vector>

namespace rethink::addin {
namespace {

constexpr GUID kDiidEApplication = {0x914934C2, 0x5A91, 0x11CF, {0x87, 0x00, 0x00, 0xAA, 0x00, 0x60, 0x26, 0x3B}};

ComPtr<IDispatch> g_app;
ComPtr<IConnectionPoint> g_point;
DWORD g_cookie = 0;
GUID g_eventIid = kDiidEApplication;
long g_eventCount = 0;
std::wstring g_eventError;
DISPID g_selectionChange = 2001, g_beforeDoubleClick = 2003;  // aus der Typbibliothek überschrieben

// Aktuell markiertes Element und zwischengespeicherte Segmente (Folienpunkte)
ComPtr<IDispatch> g_element;
std::vector<SegmentInfo> g_segments;
double g_cachedLeft = -1, g_cachedTop = -1, g_cachedWidth = -1, g_cachedHeight = -1;
OverlayContent g_shown;
UINT_PTR g_trackTimer = 0;
bool g_acting = false;

// Ziehen eines ▲▼-Griffs (Maßstab und Schrittweite zu Beginn eingefroren)
struct Drag {
    bool active = false;  // beginSegmentDrag war erfolgreich
    double start = 0, pixelsPerUnit = 0;
    int decimals = 0;
    std::wstring error;  // erst nach dem Loslassen melden: während des Ziehens hält das Overlay die Maus
};
bool g_dragging = false;
Drag g_drag;
double g_lastPixelsPerUnit = 0;  // für Tests

// Leuchtrahmen folgt dem Verschieben des Elements in PowerPoint (Maus-Hook, nur dieser Thread)
enum class Follow { None, Armed, Moving, Resizing };
HHOOK g_mouseHook = nullptr;
Follow g_follow = Follow::None;
POINT g_followFrom{};

void stopTracking() {
    g_follow = Follow::None;
    g_element.Reset();
    g_segments.clear();
    g_cachedWidth = -1;
    g_shown = {};
    hideOverlay();
    if (g_trackTimer) KillTimer(nullptr, g_trackTimer);
    g_trackTimer = 0;
}

void remember(const char* where, const std::exception& e) { g_eventError = widen(where) + L": " + widen(e.what()); }

// Hauptfenster von PowerPoint (Application.HWND ist nicht in jeder Version per IDispatch erreichbar).
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

int scaled(int px) {
    const UINT dpi = GetDpiForWindow(powerPointWindow());
    return MulDiv(px, dpi ? dpi : 96, 96);
}

// Folienpunkte → Bildschirmpixel (linear; zwei Stützstellen pro Achse)
struct Mapping {
    double ax = 1, bx = 0, ay = 1, by = 0;
    long x(double p) const { return std::lround(ax * p + bx); }
    long y(double p) const { return std::lround(ay * p + by); }
};

Mapping currentMapping() {
    auto window = disp::object(g_app.Get(), L"ActiveWindow");
    const double x0 = disp::call(window.Get(), L"PointsToScreenPixelsX", 0.0).toDouble();
    const double x1 = disp::call(window.Get(), L"PointsToScreenPixelsX", 1000.0).toDouble();
    const double y0 = disp::call(window.Get(), L"PointsToScreenPixelsY", 0.0).toDouble();
    const double y1 = disp::call(window.Get(), L"PointsToScreenPixelsY", 1000.0).toDouble();
    return {(x1 - x0) / 1000.0, x0, (y1 - y0) / 1000.0, y0};
}

// PowerPoints eigene Anfasser (8 Größenpunkte + Drehpunkt) dürfen nicht von Griffen verdeckt werden.
bool hitsGrip(const RECT& r, const RECT& e) {
    const int g = scaled(7);
    const long mx = (e.left + e.right) / 2, my = (e.top + e.bottom) / 2;
    const POINT grips[] = {{e.left, e.top}, {mx, e.top}, {e.right, e.top}, {e.left, my}, {e.right, my},
                           {e.left, e.bottom}, {mx, e.bottom}, {e.right, e.bottom}, {mx, e.top - scaled(22)}};
    for (auto& p : grips) {
        RECT grip{p.x - g, p.y - g, p.x + g, p.y + g};
        RECT overlap{};
        if (IntersectRect(&overlap, &grip, &r)) return true;
    }
    return false;
}

OverlayContent buildContent(double left, double top, double width, double height) {
    const Mapping m = currentMapping();
    OverlayContent c;
    c.element = {m.x(left), m.y(top), m.x(left + width), m.y(top + height)};
    const int half = scaled(5), tall = scaled(8), gap = scaled(3);
    for (auto& s : g_segments) {
        if (s.flags == "e" || s.flags.empty()) continue;  // berechnete Segmente haben keine Griffe
        const long cx = m.x(s.left + s.width / 2);
        const long edge = s.flags == "neg" ? m.y(s.top + s.height) : m.y(s.top);  // Ende des Segments
        RECT up{cx - half, edge - gap - tall, cx + half, edge - gap};
        RECT down{cx - half, edge + gap, cx + half, edge + gap + tall};
        if (hitsGrip(up, c.element) || hitsGrip(down, c.element)) {
            OffsetRect(&up, scaled(14), 0);
            OffsetRect(&down, scaled(14), 0);
        }
        c.handles.push_back({up, OverlayAction::Up, s.row, s.column});
        c.handles.push_back({down, OverlayAction::Down, s.row, s.column});
    }
    const int icon = scaled(16), offset = scaled(9);
    c.handles.push_back({{c.element.right + offset, c.element.bottom + offset, c.element.right + offset + icon, c.element.bottom + offset + icon},
                         OverlayAction::Datasheet, 0, 0});
    return c;
}

bool sameContent(const OverlayContent& a, const OverlayContent& b) {
    if (!EqualRect(&a.element, &b.element) || a.handles.size() != b.handles.size()) return false;
    for (size_t i = 0; i < a.handles.size(); ++i)
        if (!EqualRect(&a.handles[i].rect, &b.handles[i].rect)) return false;
    return true;
}

void onOverlayAction(const OverlayHandle& handle, OverlayGesture gesture, int pixels, bool shift);

// Overlay an das Element anpassen (Zoom, Bildlauf, Verschieben, Größe). force = Segmente neu einlesen.
void placeOverlay(bool force) {
    const double left = disp::call(g_element.Get(), L"Left").toDouble(), top = disp::call(g_element.Get(), L"Top").toDouble();
    const double width = disp::call(g_element.Get(), L"Width").toDouble(), height = disp::call(g_element.Get(), L"Height").toDouble();
    if (force || width != g_cachedWidth || height != g_cachedHeight || left != g_cachedLeft || top != g_cachedTop) {
        g_segments = segmentsOf(g_element.Get());
        g_cachedLeft = left, g_cachedTop = top, g_cachedWidth = width, g_cachedHeight = height;
    }
    const OverlayContent content = buildContent(left, top, width, height);
    if (!force && overlayVisible() && !overlayDisplaced() && sameContent(content, g_shown)) return;
    g_shown = content;
    showOverlay(powerPointWindow(), content, onOverlayAction);
}

// Overlay während des Ziehens: Segmente kommen aus der Zeichnung, nicht aus erneuten Abfragen der Formen.
void placeOverlayFor(const std::vector<SegmentInfo>& segments) {
    if (!g_element) return;
    g_segments = segments;
    g_cachedLeft = disp::call(g_element.Get(), L"Left").toDouble(), g_cachedTop = disp::call(g_element.Get(), L"Top").toDouble();
    g_cachedWidth = disp::call(g_element.Get(), L"Width").toDouble(), g_cachedHeight = disp::call(g_element.Get(), L"Height").toDouble();
    g_shown = buildContent(g_cachedLeft, g_cachedTop, g_cachedWidth, g_cachedHeight);
    showOverlay(powerPointWindow(), g_shown, onOverlayAction);
}

void track() {
    auto current = rethinkElementOf(disp::object(disp::object(g_app.Get(), L"ActiveWindow").Get(), L"Selection").Get());
    if (!current) return stopTracking();
    g_element = current;
    // Größe geändert? Erst nach dem Loslassen der Maus neu zeichnen.
    if ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) == 0 && relayoutIfResized(g_element.Get())) {
        placeOverlay(true);
        return;
    }
    placeOverlay(false);
}

void CALLBACK onTrack(HWND, UINT, UINT_PTR, DWORD) {
    if (!g_element || g_acting || g_dragging) return;
    if (g_follow != Follow::None) {
        if (GetAsyncKeyState(VK_LBUTTON) & 0x8000) return;  // PowerPoint verschiebt noch oder ändert die Größe
        g_follow = Follow::None;                             // Loslassen verpasst: ab jetzt wieder normal nachführen
    }
    try {
        track();
    } catch (const std::exception& e) {
        remember("Nachführen", e);
        stopTracking();
    } catch (...) {
        stopTracking();
    }
}

// Eine Geste des Ziehens verarbeiten. Liefert bei DragEnd/DragCancel einen aufgelaufenen Fehler (sonst leer).
std::wstring dragGesture(const OverlayHandle& handle, OverlayGesture gesture, int pixels) {
    const bool finish = gesture == OverlayGesture::DragEnd || gesture == OverlayGesture::DragCancel;
    try {
        if (!g_dragging) {
            if (gesture == OverlayGesture::DragCancel) return {};  // abgebrochen, bevor etwas geändert wurde
            g_dragging = true;
            g_drag = {};
            const SegmentDragStart start = beginSegmentDrag(g_app.Get(), handle.row, handle.column);
            g_drag.start = start.value;
            g_drag.decimals = start.decimals;
            g_drag.pixelsPerUnit = start.pointsPerUnit * std::fabs(currentMapping().ay);
            // Alle Segmente 0: kein Maßstab ablesbar – dann 4 Pixel je Schritt
            if (g_drag.pixelsPerUnit <= 0) g_drag.pixelsPerUnit = scaled(4) * std::pow(10.0, start.decimals);
            g_lastPixelsPerUnit = g_drag.pixelsPerUnit;
            g_drag.active = true;
        }
        if (g_drag.active) {
            const double value = gesture == OverlayGesture::DragCancel
                                     ? g_drag.start
                                     : engine::dragValue(g_drag.start, pixels, g_drag.pixelsPerUnit, g_drag.decimals);
            std::vector<SegmentInfo> segments;
            if (dragSegmentTo(value, segments)) {
                try {
                    placeOverlayFor(segments);
                } catch (...) {
                    // Griffe bleiben stehen; nach dem Loslassen setzt track() das Overlay neu
                }
            }
        }
    } catch (const std::exception& e) {
        if (g_drag.error.empty()) g_drag.error = widen(e.what());
        g_drag.active = false;
    } catch (const winrt::hresult_error& e) {
        if (g_drag.error.empty()) g_drag.error = e.message().c_str();
        g_drag.active = false;
    } catch (...) {
        if (g_drag.error.empty()) g_drag.error = L"Unerwarteter Fehler beim Ziehen";
        g_drag.active = false;
    }
    if (!finish) return {};

    std::wstring error = std::move(g_drag.error);
    try {
        endSegmentDrag();
    } catch (const std::exception& e) {
        if (error.empty()) error = widen(e.what());
    } catch (...) {
    }
    g_dragging = false;
    g_drag = {};
    try {
        if (g_element) track();  // Element kann bei einem vollständigen Neuzeichnen ersetzt worden sein
    } catch (...) {
        stopTracking();
    }
    return error;
}

void onOverlayAction(const OverlayHandle& handle, OverlayGesture gesture, int pixels, bool shift) {
    if (!g_app) return;
    if (gesture != OverlayGesture::Click) {
        const std::wstring error = dragGesture(handle, gesture, pixels);
        if (!error.empty()) MessageBoxW(powerPointWindow(), error.c_str(), L"rethink", MB_ICONINFORMATION);
        return;
    }
    if (g_acting || g_dragging) return;
    g_acting = true;
    try {
        if (handle.action == OverlayAction::Datasheet) {
            commandOpenDatasheet(g_app.Get());
        } else {
            adjustSegment(g_app.Get(), handle.row, handle.column, handle.action == OverlayAction::Up ? 1 : -1, shift);
            if (g_element) placeOverlay(true);
        }
    } catch (const std::exception& e) {
        MessageBoxW(powerPointWindow(), widen(e.what()).c_str(), L"rethink", MB_ICONINFORMATION);
    } catch (const winrt::hresult_error& e) {
        MessageBoxW(powerPointWindow(), e.message().c_str(), L"rethink", MB_ICONINFORMATION);
    } catch (...) {
    }
    g_acting = false;
}

// Maus-Hook: PowerPoint meldet Left/Top erst beim Loslassen. Bis dahin wird das Overlay um den Mausweg versetzt;
// beim Ändern der Größe (Drücken auf einen Anfasser) verschwindet es und kommt nach dem Neu-Layout wieder.
LRESULT CALLBACK onMouse(int code, WPARAM message, LPARAM lParam) {
    if (code >= 0 && g_element && !g_dragging && !g_acting) {
        const auto* info = reinterpret_cast<const MOUSEHOOKSTRUCT*>(lParam);
        if (info->hwnd != overlayWindow()) {
            const POINT p = info->pt;
            switch (message) {
                case WM_LBUTTONDOWN: {
                    g_follow = Follow::None;
                    if (!overlayVisible()) break;
                    const RECT probe{p.x, p.y, p.x + 1, p.y + 1};
                    if (hitsGrip(probe, g_shown.element)) g_follow = Follow::Resizing;
                    else if (PtInRect(&g_shown.element, p)) g_follow = Follow::Armed;
                    g_followFrom = p;
                    break;
                }
                case WM_MOUSEMOVE: {
                    if (g_follow == Follow::None || (GetAsyncKeyState(VK_LBUTTON) & 0x8000) == 0) break;
                    const int dx = p.x - g_followFrom.x, dy = p.y - g_followFrom.y;
                    if (g_follow == Follow::Armed && std::abs(dx) < GetSystemMetrics(SM_CXDRAG) && std::abs(dy) < GetSystemMetrics(SM_CYDRAG))
                        break;
                    if (g_follow == Follow::Resizing) {
                        if (overlayVisible()) hideOverlay();
                    } else {
                        g_follow = Follow::Moving;
                        offsetOverlay(dx, dy);
                    }
                    break;
                }
                case WM_LBUTTONUP: g_follow = Follow::None; break;  // der nächste Takt setzt das Overlay an die neue Lage
            }
        }
    }
    return CallNextHookEx(g_mouseHook, code, message, lParam);
}
void onSelectionChange(IDispatch* selection) {
    try {
        auto element = rethinkElementOf(selection);
        if (!element) return stopTracking();
        g_element = element;
        placeOverlay(true);
        if (!g_trackTimer) g_trackTimer = SetTimer(nullptr, 0, 150, onTrack);
    } catch (const std::exception& e) {
        remember("Auswahl", e);
        stopTracking();
    } catch (...) {
        stopTracking();
    }
}

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
            if (g_acting || g_dragging) return S_OK;  // eigene Änderungen lösen keine Neuberechnung des Overlays aus
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
    sink->Release();
    g_mouseHook = SetWindowsHookExW(WH_MOUSE, onMouse, nullptr, GetCurrentThreadId());
}

void stopEvents() {
    if (g_mouseHook) UnhookWindowsHookEx(g_mouseHook);
    g_mouseHook = nullptr;
    if (g_dragging) {
        try {
            endSegmentDrag();
        } catch (...) {
        }
        g_dragging = false;
    }
    stopTracking();
    if (g_point && g_cookie) g_point->Unadvise(g_cookie);
    g_point.Reset();
    g_cookie = 0;
    g_app.Reset();
}

std::wstring eventStatus() {
    wchar_t iid[64] = {};
    StringFromGUID2(g_eventIid, iid, 64);
    return std::format(L"verbunden={} iid={} auswahl={} doppelklick={} ereignisse={} overlay={} datenblaetter={}{}", g_cookie != 0, iid,
                       g_selectionChange, g_beforeDoubleClick, g_eventCount, overlayVisible(), watchCount(),
                       g_eventError.empty() ? L"" : L" fehler=" + g_eventError);
}

std::wstring overlayInfo() {
    const RECT r = overlayRect();
    std::wstring text = std::format(L"sichtbar={} fenster={},{},{},{} element={},{},{},{} griffe={}", overlayVisible(), r.left, r.top, r.right,
                                    r.bottom, g_shown.element.left, g_shown.element.top, g_shown.element.right, g_shown.element.bottom,
                                    g_shown.handles.size());
    for (auto& h : g_shown.handles)
        text += std::format(L" [{} r{} c{} {},{},{},{}]",
                            h.action == OverlayAction::Up ? L"auf" : h.action == OverlayAction::Down ? L"ab" : L"datenblatt", h.row, h.column,
                            h.rect.left, h.rect.top, h.rect.right, h.rect.bottom);
    return text;
}

std::wstring simulateSegmentDrag(int row, int column, const std::vector<int>& pixels) {
    if (!g_app) throw Error("Das Add-in ist nicht mit PowerPoint verbunden");
    if (g_dragging) throw Error("Es wird bereits gezogen");
    const OverlayHandle handle{{}, OverlayAction::Up, row, column};
    for (int p : pixels) dragGesture(handle, OverlayGesture::Drag, p);
    const std::wstring error = dragGesture(handle, OverlayGesture::DragEnd, pixels.empty() ? 0 : pixels.back());
    if (!error.empty()) throw Error(narrow(error));
    return std::format(L"pixelProEinheit={:.4f}", g_lastPixelsPerUnit);
}

}  // namespace rethink::addin
