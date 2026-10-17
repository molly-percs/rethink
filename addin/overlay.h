#pragma once

#include <windows.h>

#include <string>

namespace rethink::addin {

// Durchsichtiges, nicht anklickbares Fenster über PowerPoint, das ein ausgewähltes rethink-Element
// mit einem blauen Leuchtrahmen und einem kleinen Etikett markiert. Das Dokument bleibt unverändert.
// `rect` in Bildschirmpixeln; `owner` ist das PowerPoint-Hauptfenster (das Overlay bleibt darüber).
void showOverlay(HWND owner, const RECT& rect, const std::wstring& label);
void hideOverlay();
bool overlayVisible();

}  // namespace rethink::addin
