#pragma once

#include "util.h"

#include <filesystem>
#include <string>
#include <vector>

namespace rethink::addin {

// --- Befehle ohne Dialoge (auch über COMAddIn.Object per Automation aufrufbar) ---

// Erzeugt aus `input` eine Präsentation und öffnet sie. Ist `output` leer, entsteht
// <input>.pptx; existiert die Datei schon oder ist sie geöffnet, ein freier Name "<input> (2).pptx".
// Liefert den Pfad der geöffneten Präsentation.
std::wstring buildAndOpen(IDispatch* app, const std::filesystem::path& input, std::filesystem::path output,
                          std::vector<std::string>& warnings);

// Fügt auf der aktuellen Folie ein gruppiertes Säulendiagramm ein und benennt es "Chart<n>".
// Liefert den vergebenen Namen.
std::wstring insertChart(IDispatch* app);

// Setzt den Alternativtext "rethink:<name>" auf alle ausgewählten Formen; ein leerer Name entfernt ihn.
// Liefert die Zahl der geänderten Formen.
int setElementName(IDispatch* app, const std::wstring& name);

// rethink-Name der ersten ausgewählten Form (aus dem Alternativtext, sonst der Formname).
std::wstring elementName(IDispatch* app);

// --- Ribbon-Befehle: mit Dialogen, Fehler werden als Meldung angezeigt ---

void commandBuild(IDispatch* app);
void commandInsertChart(IDispatch* app);
void commandSetName(IDispatch* app);

}  // namespace rethink::addin
