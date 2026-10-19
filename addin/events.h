#pragma once

#include "elements.h"
#include "util.h"

#include <string>
#include <vector>

namespace rethink::addin {

// Abonniert PowerPoint-Ereignisse (EApplication):
//  - Auswahlwechsel: ein ausgewähltes rethink-Element erhält den blauen Leuchtrahmen und seine Griffe (overlay.h);
//  - Doppelklick auf ein rethink-Element öffnet sein Datenblatt statt der Gruppenbearbeitung (Zeitplan: auf der Skala
//    den Kalender des Bereichs, auf einem Element dessen Kalender).
void startEvents(IDispatch* app);
void stopEvents();

// Diagnose für Tests: Verbindung, Ereignis-IDs, Anzahl empfangener Ereignisse, Overlay sichtbar.
std::wstring eventStatus();

// Diagnose für Tests: Overlay-Fenster, Elementrechteck und alle Griffe in Bildschirmpixeln.
std::wstring overlayInfo();

// Für Tests: zieht den Wertgriff des Segments (row, column) des ausgewählten Elements über dieselben Schritte wie die
// Maus (pixels: Mausweg je Bewegung, positiv = Richtung wachsender Werte) und lässt los. Liefert den eingefrorenen Maßstab.
std::wstring simulateSegmentDrag(int row, int column, const std::vector<int>& pixels);

// Für Tests: zieht am ausgewählten Zeitplan (role wie "bar r2", "item 0" oder "range"; grip 0 = Anfang, 1 = Ende,
// 2 = verschieben; "row rN" zieht Zeile N senkrecht, Pixel nach unten) über dieselben Schritte wie die Maus (Pixel nach rechts) und lässt los bzw. bricht ab (Esc).
// Liefert die Kurzinfos der einzelnen Schritte.
std::wstring simulateGanttDrag(const std::string& role, int grip, const std::vector<int>& pixels, bool cancel);

// Was liegt im ausgewählten Zeitplan unter dem Mauszeiger (Kontextmenü, Doppelklick) bzw. an einem Folienpunkt (Tests)?
GanttHit ganttHitAtCursor(IDispatch* app);
GanttHit ganttHitAtPoint(IDispatch* app, double x, double y);
// Für Tests: Rechtsklick an einem Folienpunkt des ausgewählten Zeitplans. Freie Stelle im Zeitbereich: öffnet das Rundmenü,
// liefert seine Schaltflächen und löst command aus (leer = offen lassen); sonst leer (dort käme das Kontextmenü).
std::wstring simulateGanttRightClick(double x, double y, const std::wstring& command);
// Für Tests: Maus über einem Folienpunkt des ausgewählten Zeitplans. Liefert die hervorgehobene Grenze ("waagerecht zeile=3 band=…",
// "senkrecht datum=2026-11-09 band=…") und den leuchtenden Griff ("griff=zelle r2 c0"), durch "; " getrennt; leer = nichts.
std::wstring simulateGanttHover(double x, double y);
// Für Tests: öffnet das Eingabefeld der Beschriftungszelle an einem Folienpunkt wie ein Klick, setzt den Text, bestätigt wie Enter
// und liefert den Zelltext danach. Beginnt text mit U+0001, bleibt das Eingabefeld mit dem restlichen Text offen (Zelltext davor).
std::wstring simulateGanttEditCell(double x, double y, const std::wstring& text);

}  // namespace rethink::addin
