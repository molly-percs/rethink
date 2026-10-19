#pragma once

#include <windows.h>

#include <string>
#include <vector>

namespace rethink::addin {

// Rundmenü für neue Elemente (Rechtsklick auf eine freie Stelle im Zeitbereich des Zeitplans): runde Schaltflächen
// im Kreis um den Mauspunkt, in der Mitte „schließen“, dazu ein rotes Fadenkreuz über dem Zeitbereich an der Stelle,
// an der das neue Element entsteht. Eigenes Ebenenfenster; Esc oder ein Klick daneben schließt es.

struct RadialButton {
    std::wstring command;  // z. B. "new=bar"; das Symbol richtet sich nach dem Befehl
    std::wstring label;    // Kurzinfo beim Darüberfahren
    bool enabled = true;
};

// Wird nach einem Klick auf eine Schaltfläche aufgerufen (das Menü ist dann schon zu).
using RadialCallback = void (*)(const std::wstring& command);

// plot: Zeitbereich in Bildschirmpixeln (für das Fadenkreuz).
void showRadialMenu(HWND owner, POINT center, const RECT& plot, std::vector<RadialButton> buttons, RadialCallback callback);
void hideRadialMenu();
HWND radialMenuWindow();  // nullptr = zu

// Für Tests: offene Schaltflächen ("Beschriftung=Befehl", "(aus)" wenn gesperrt) bzw. eine davon wie per Klick auslösen.
std::wstring radialMenuText();
bool clickRadialButton(const std::wstring& command);

}  // namespace rethink::addin
