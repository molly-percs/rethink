#pragma once

#include "elements.h"

#include <windows.h>

#include <functional>

namespace rethink::addin {

// Einfügemenü wie im Original: ein Raster aller Diagrammtypen; noch nicht umgesetzte Typen sind ausgegraut.
// Bei Typen mit Ausrichtung erscheinen beim Überfahren Pfeile um das Symbol – ein Klick darauf fügt in dieser
// Ausrichtung ein, ein Klick auf das Symbol in der Grundausrichtung. Esc oder ein Klick daneben schließt.
struct GalleryChoice {
    ElementKind kind = ElementKind::Waterfall;
    engine::Orientation orientation = engine::Orientation::Up;
};

void openGallery(HWND owner, POINT at, std::function<void(GalleryChoice)> chosen);
bool galleryOpen();

}  // namespace rethink::addin
