#pragma once

#include "util.h"

#include <msopc.h>

#include <filesystem>
#include <map>

namespace rethink {

using PartUri = ComPtr<IOpcPartUri>;

struct Relationship {
    std::wstring id;
    std::wstring type;
    bool external = false;
    ComPtr<IUri> targetUri;  // wie in der .rels-Datei angegeben
    PartUri targetPart;      // aufgelöst; nur bei internen Zielen, die gültige Teilnamen sind
};

// Ein OPC-Paket (PPTX, XLSX) über die Packaging API des Windows SDK.
class Package {
public:
    // Quell-Teilname (klein geschrieben) → Kopie; verhindert Doppelkopien innerhalb einer Folie.
    using CopyMemo = std::map<std::wstring, PartUri>;

    static Package open(const std::filesystem::path& file);
    static Package openBytes(const std::vector<std::uint8_t>& bytes);
    static Package createEmpty();
    void save(const std::filesystem::path& file) const;
    std::vector<std::uint8_t> saveToBytes() const;

    PartUri uri(std::wstring_view partName) const;
    static std::wstring name(IOpcUri* uri);
    bool exists(IOpcPartUri* uri) const;
    ComPtr<IOpcPart> part(IOpcPartUri* uri) const;
    ComPtr<IOpcPart> createPart(std::wstring_view partName, std::wstring_view contentType);
    ComPtr<IStream> stream(IOpcPart* part) const;
    std::vector<std::uint8_t> read(IOpcPart* part) const;
    void write(IOpcPart* part, const std::vector<std::uint8_t>& bytes) const;
    void writeText(IOpcPart* part, std::wstring_view xml) const;
    // Freier Teilname nach dem Muster von `like`, z. B. /ppt/charts/chart7.xml.
    std::wstring uniqueName(std::wstring_view like) const;

    // source == nullptr: Beziehungen des Pakets selbst.
    std::vector<Relationship> relationships(IOpcPart* source) const;
    std::wstring addRelationship(IOpcPart* source, std::wstring_view type, IOpcPartUri* target, std::wstring_view id = {});
    void deleteRelationship(IOpcPart* source, std::wstring_view id);

    // Kopiert einen Teil samt eigener Unterteile (Diagramme, Einbettungen, Tags).
    // Layouts, Master, Designs und Medien werden geteilt, Notizen entfallen.
    PartUri copyPartDeep(IOpcPartUri* source, CopyMemo& memo);
    // Entfernt Teile, die von der Paketwurzel aus nicht mehr erreichbar sind.
    size_t collectGarbage();

private:
    void init();
    ComPtr<IOpcRelationshipSet> relationshipSet(IOpcPart* source) const;
    ComPtr<IOpcUri> sourceUri(IOpcPart* source) const;

    ComPtr<IOpcFactory> factory_;
    ComPtr<IOpcPackage> package_;
    ComPtr<IOpcPartSet> parts_;
};

}  // namespace rethink
