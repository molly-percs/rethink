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

// Wird nach einem Klick auf einen Griff aufgerufen (nicht innerhalb der Mausbehandlung, sondern nachgereicht).
using OverlayCallback = void (*)(const OverlayHandle& handle, bool shift);

void showOverlay(HWND owner, const OverlayContent& content, OverlayCallback callback);
void hideOverlay();
bool overlayVisible();
// Fensterrechteck des Overlays (für Tests).
RECT overlayRect();

}  // namespace rethink::addin
