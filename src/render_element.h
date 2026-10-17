#pragma once

#include "opc.h"
#include "ppttc.h"
#include "xmldoc.h"

#include <string>
#include <vector>

namespace rethink {

inline constexpr wchar_t kModelTag[] = L"RETHINK_MODEL";

// Rendert gezeichnete rethink-Diagramme auf einer Folie (Layout-Kern + OOXML-Ausgabe):
//  - bestehende rethink-Elemente (Gruppe mit Tag RETHINK_MODEL) namens item.name werden mit neuen Daten neu gezeichnet;
//    ein "chart"-Objekt im Eintrag überschreibt einzelne gespeicherte Optionen;
//  - hat der Eintrag ein "chart"-Objekt, wird eine gleichnamige Platzhalterform durch das Diagramm ersetzt.
// Gibt die Anzahl gerenderter Elemente zurück.
int renderChartElements(Package& pkg, IOpcPart* slidePart, const XmlDoc& slide, std::vector<Relationship>& slideRels,
                        const DataItem& item, std::vector<std::string>& warnings);

}  // namespace rethink
