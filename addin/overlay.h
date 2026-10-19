#pragma once

#include <windows.h>

#include <string>
#include <vector>

namespace rethink::addin {

// Überlagerung eines ausgewählten rethink-Elements (eigenes Fenster über PowerPoint, Dokument bleibt unverändert):
// blauer Leuchtrahmen, winziges Etikett „rethink chart“, Griffe und ein Datenblatt-Symbol unten rechts.
// Nur Griffe und Symbol nehmen Klicks an; alles andere reicht Klicks an PowerPoint durch.

enum class OverlayAction {
    Up = 1,         // Wert erhöhen (▲ bzw. in Richtung wachsender Werte)
    Down = 2,       // Wert verringern
    Datasheet = 3,  // Datenblatt-Symbol
    GanttStart,     // Zeitplan: Anfang eines Balkens/einer Klammer/Schattierung ziehen
    GanttEnd,       // … Ende ziehen
    GanttMove,      // … ganzes Element verschieben (Klick wählt die Form aus)
    RangeStart,     // Anfang des sichtbaren Bereichs (Pfeil am linken Skalenende)
    RangeEnd,       // Ende des sichtbaren Bereichs
    RowMove,        // Zeitplan: Zeile am Griff links neben den Beschriftungen nach oben/unten ziehen (auch in eine andere Kategorie)
    CellEdit,       // Zeitplan: Zelle einer Beschriftungsspalte (Klick = dort tippen)
};

enum class HandleStyle {
    Triangle,   // ▲▼◀▶ an Segmenten
    Grip,       // schmaler Griff an einem Balkenende
    Hidden,     // unsichtbare Fläche (Balken/Meilenstein verschieben)
    Arrow,      // Pfeil am Skalenende
    DoubleArrow,// ◀▶ am Ende des blauen Skalenbands (Skala ausgewählt oder Bereich wird gezogen)
    Datasheet,  // Tabellensymbol
    RowGrip,    // Zeilengriff „⋮⋮“ (blass, bei Maus über der Zeile kräftig)
};

struct OverlayHandle {
    RECT rect{};  // Bildschirmpixel
    OverlayAction action = OverlayAction::Up;
    int row = 0, column = 0;
    std::string role;  // Zeitplan: Rolle der Form ("bar r3", "item 2" …)
    HandleStyle style = HandleStyle::Triangle;
    int dirX = 0, dirY = -1;  // Bildschirmrichtung, in der der Mausweg positiv zählt
    int pointing = 0;         // Dreieck/Pfeil: 0 = oben, 1 = unten, 2 = links, 3 = rechts
    RECT area{};              // Zeile/Zelle: Beschriftungsbereich der Zeile (heller Streifen bei Maus darüber)
    RECT cell{};              // Zelle: volle Zelle für das Leuchten (rect kann schmaler sein, um PowerPoints Anfasser freizulassen)
};

struct OverlayContent {
    RECT element{};  // Bildschirmpixel
    std::vector<OverlayHandle> handles;
    std::wstring tooltip;  // Kurzinfo beim Ziehen (z. B. Datum)
    POINT tooltipAt{};     // Bildschirmpunkt (Mitte unten der Kurzinfo)
    RECT marker{};         // blaue Einfügelinie beim Ziehen einer Zeile (leer = keine)
    RECT band{};           // blaues Band über der Skala (leer = keins)
    RECT line{};           // helles blaues Band an der Grenze unter der Maus (Trennlinie einfügen; leer = keins)
};

// Geste an einem Griff. Ein Griff wird geklickt (Loslassen ohne Bewegung) oder gezogen: Drag kommt bei jeder
// Bewegung (zusammengefasst, nur der neueste Stand), danach genau einmal DragEnd oder DragCancel (Esc, Fokusverlust).
enum class OverlayGesture { Click, Drag, DragEnd, DragCancel };

// Wird nach einer Geste aufgerufen (nicht innerhalb der Mausbehandlung, sondern nachgereicht). pixels: Mausweg seit
// dem Drücken in Bildschirmpixeln entlang (dirX, dirY) des Griffs. Bei DragEnd/DragCancel ist die Maus frei.
using OverlayCallback = void (*)(const OverlayHandle& handle, OverlayGesture gesture, int pixels, bool shift);

void showOverlay(HWND owner, const OverlayContent& content, OverlayCallback callback);
void hideOverlay();
bool overlayVisible();
// Prüft, ob die Maus noch über demselben Griff liegt (Maus-Hook meldet Bewegungen außerhalb des Overlays) und zeichnet bei Änderung neu.
void overlayHoverCheck();
// Setzt das helle Band an der Grenze unter der Maus (Bildschirmpixel, leer = keins) und zeichnet nur bei Änderung neu.
void setOverlayLine(const RECT& line);
// Für Tests: tut so, als läge die Maus bei `screen`; liefert den Griff dort (gültig bis zum nächsten Aufbau) oder nullptr.
const OverlayHandle* overlayHoverAt(POINT screen);
// Fensterrechteck des Overlays (für Tests).
RECT overlayRect();
HWND overlayWindow();

// Verschiebt das sichtbare Overlay, ohne es neu zu zeichnen (Leuchtrahmen folgt dem Verschieben in PowerPoint,
// das Left/Top erst beim Loslassen meldet). dx/dy relativ zur zuletzt gezeigten Lage.
void offsetOverlay(int dx, int dy);
// Steht das Overlay nicht mehr dort, wo showOverlay es hingesetzt hat?
bool overlayDisplaced();

}  // namespace rethink::addin
