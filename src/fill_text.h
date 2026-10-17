#pragma once

#include "ppttc.h"
#include "xmldoc.h"

#include <string>
#include <vector>

namespace rethink {

// Ein Element heißt `name`, wenn der Formname (Auswahlbereich) gleich ist
// oder der Alternativtext "rethink:<name>" lautet.
bool matchesName(const XmlDoc& doc, IXMLDOMNode* cNvPr, const std::wstring& name);

// Füllt alle Textziele dieser Folie, die `item` benennt:
//  - think-cell-kompatible Felder (a:fld type="thinkcell…<m_strName>NAME</m_strName>…")
//  - benannte Formen mit Textkörper
// Gibt die Anzahl gefüllter Ziele zurück.
int fillTextInSlide(const XmlDoc& slide, const DataItem& item, std::vector<std::string>& warnings);

}  // namespace rethink
