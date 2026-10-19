#pragma once

#include <ocidl.h>

#include <string>

namespace rethink::addin {

// Eigene Symbole des Ribbon-Tabs (GDI+, kantengeglättet, mit Alphakanal), passend zu getImage im Ribbon-XML.
// id: Symbolname (siehe Tabelle in ribbon_icons.cpp); large: 32 px statt 16 px, bei höherer DPI entsprechend mehr Pixel.
// Liefert ein neues IPictureDisp (der Aufrufer übernimmt die Referenz) oder nullptr bei unbekanntem id.
IPictureDisp* ribbonIcon(const std::wstring& id, bool large);

}  // namespace rethink::addin
