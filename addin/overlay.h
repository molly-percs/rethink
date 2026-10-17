#pragma once

#include <windows.h>

#include <string>
#include <vector>

namespace rethink::addin {

// Überlagerung eines ausgewählten rethink-Elements (eigenes Fenster über PowerPoint, Dokument bleibt unverändert):
// blauer Leuchtrahmen, winziges Etikett „rethink chart“, ▲▼-Griffe an den Segmenten und ein Datenblatt-Symbol
// unten rechts. Nur Griffe und Symbol nehmen Klicks an; alles andere reicht Klicks an PowerPoint durch.

enum class OverlayAction { Up = 1, Down = 2, Datasheet = 3 };

struct OverlayHandle {
    RECT rect{};  // Bildschirmpixel
    OverlayAction action = OverlayAction::Up;
    int row = 0, column = 0;
};

struct OverlayContent {
    RECT element{};  // Bildschirmpixel
    std::vector<OverlayHandle> handles;
};

// Geste an einem Griff. Ein ▲▼-Griff wird geklickt (Loslassen ohne Bewegung) oder gezogen: Drag kommt bei jeder
// Bewegung (zusammengefasst, nur der neueste Stand), danach genau einmal DragEnd oder DragCancel (Esc, Fokusverlust).
enum class OverlayGesture { Click, Drag, DragEnd, DragCancel };

// Wird nach einer Geste aufgerufen (nicht innerhalb der Mausbehandlung, sondern nachgereicht). pixels: senkrechter
// Mausweg seit dem Drücken in Bildschirmpixeln, positiv = nach oben. Bei DragEnd/DragCancel ist die Maus frei.
using OverlayCallback = void (*)(const OverlayHandle& handle, OverlayGesture gesture, int pixels, bool shift);

void showOverlay(HWND owner, const OverlayContent& content, OverlayCallback callback);
void hideOverlay();
bool overlayVisible();
// Fensterrechteck des Overlays (für Tests).
RECT overlayRect();
HWND overlayWindow();

// Verschiebt das sichtbare Overlay, ohne es neu zu zeichnen (Leuchtrahmen folgt dem Verschieben in PowerPoint,
// das Left/Top erst beim Loslassen meldet). dx/dy relativ zur zuletzt gezeigten Lage.
void offsetOverlay(int dx, int dy);
// Steht das Overlay nicht mehr dort, wo showOverlay es hingesetzt hat?
bool overlayDisplaced();

}  // namespace rethink::addin
