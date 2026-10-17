#pragma once

#include <filesystem>
#include <string>
#include <vector>

namespace rethink {

struct BuildResult {
    int slides = 0;          // Folien in der Ausgabe
    int filled = 0;          // gefüllte Elemente (Textziele + Diagramme)
    size_t removedParts = 0; // nicht mehr referenzierte Teile, die entfernt wurden
    std::vector<std::string> warnings;
};

// Erzeugt aus einer .ppttc-Datei eine neue Präsentation.
// Jeder Eintrag der Datei ergibt eine Kopie aller Folien der Vorlage.
BuildResult build(const std::filesystem::path& input, const std::filesystem::path& output);

}  // namespace rethink
