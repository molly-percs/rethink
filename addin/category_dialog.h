#pragma once

#include <windows.h>

#include <optional>
#include <string>
#include <vector>

namespace rethink::addin {

// Reihenfolge der Kategorien wählen (Liste mit „Nach oben“/„Nach unten“). Modal; liefert die gewählte Reihenfolge
// oder nichts bei Abbruch.
std::optional<std::vector<std::wstring>> chooseCategoryOrder(HWND owner, const std::vector<std::wstring>& categories);

}  // namespace rethink::addin
