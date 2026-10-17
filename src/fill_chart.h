#pragma once

#include "opc.h"
#include "ppttc.h"

#include <string>
#include <vector>

namespace rethink {

// Schreibt die Datentabelle in ein natives PowerPoint-Diagramm (c:chartSpace):
//  Zeile 0 = Kategorien (erste Zelle ignoriert), weitere Zeilen = Reihen (erste Zelle = Name).
//  Eine leere Zeile [] überspringt eine Farbe des Füllschemas, wie in think-cell dokumentiert.
// Ersetzt außerdem die eingebettete Arbeitsmappe, damit "Daten bearbeiten" dieselben Werte zeigt.
void fillChart(Package& pkg, IOpcPart* chartPart, const DataItem& item, std::vector<std::string>& warnings);

}  // namespace rethink
