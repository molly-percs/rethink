#include "opc.h"

#include <cwctype>
#include <set>

namespace rethink {
namespace {

enum class CopyRule { Copy, Share, Drop };

// Entscheidet anhand des Beziehungstyps, ob ein Ziel beim Duplizieren einer Folie
// mitkopiert, geteilt oder weggelassen wird.
CopyRule ruleFor(std::wstring_view type) {
    const auto slash = type.find_last_of(L'/');
    const std::wstring_view kind = slash == std::wstring_view::npos ? type : type.substr(slash + 1);
    static const std::set<std::wstring_view> shared = {
        L"slideLayout", L"slideMaster", L"notesMaster", L"handoutMaster", L"theme", L"image", L"audio", L"video",
        L"media", L"hdphoto", L"hyperlink", L"slide", L"font", L"presProps", L"viewProps", L"tableStyles",
        L"commentAuthors",
    };
    if (kind == L"notesSlide") return CopyRule::Drop;
    return shared.contains(kind) ? CopyRule::Share : CopyRule::Copy;
}

}  // namespace

void Package::init() {
    check(CoCreateInstance(__uuidof(OpcFactory), nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&factory_)),
          "OPC-Factory erzeugen");
}

Package Package::open(const std::filesystem::path& file) {
    Package p;
    p.init();
    ComPtr<IStream> input;
    check(p.factory_->CreateStreamOnFile(file.c_str(), OPC_STREAM_IO_READ, nullptr, FILE_ATTRIBUTE_NORMAL, &input),
          "Datei öffnen: " + narrow(file.wstring()));
    check(p.factory_->ReadPackageFromStream(input.Get(), OPC_READ_DEFAULT, &p.package_),
          "Datei als OPC-Paket lesen: " + narrow(file.wstring()));
    check(p.package_->GetPartSet(&p.parts_), "Teile des Pakets lesen");
    return p;
}

Package Package::openBytes(const std::vector<std::uint8_t>& bytes) {
    Package p;
    p.init();
    ComPtr<IStream> input;
    check(CreateStreamOnHGlobal(nullptr, TRUE, &input), "Speicher-Stream anlegen");
    writeStream(input.Get(), bytes);
    LARGE_INTEGER zero{};
    check(input->Seek(zero, STREAM_SEEK_SET, nullptr), "Stream positionieren");
    check(p.factory_->ReadPackageFromStream(input.Get(), OPC_READ_DEFAULT, &p.package_), "Daten als OPC-Paket lesen");
    check(p.package_->GetPartSet(&p.parts_), "Teile des Pakets lesen");
    return p;
}

Package Package::createEmpty() {
    Package p;
    p.init();
    check(p.factory_->CreatePackage(&p.package_), "Leeres Paket erzeugen");
    check(p.package_->GetPartSet(&p.parts_), "Teile des Pakets lesen");
    return p;
}

void Package::save(const std::filesystem::path& file) const {
    ComPtr<IStream> output;
    check(factory_->CreateStreamOnFile(file.c_str(), OPC_STREAM_IO_WRITE, nullptr, FILE_ATTRIBUTE_NORMAL, &output),
          "Ausgabedatei anlegen: " + narrow(file.wstring()));
    check(factory_->WritePackageToStream(package_.Get(), OPC_WRITE_FORCE_ZIP32, output.Get()), "Paket schreiben");
}

std::vector<std::uint8_t> Package::saveToBytes() const {
    ComPtr<IStream> memory;
    check(CreateStreamOnHGlobal(nullptr, TRUE, &memory), "Speicher-Stream anlegen");
    check(factory_->WritePackageToStream(package_.Get(), OPC_WRITE_FORCE_ZIP32, memory.Get()), "Paket schreiben");
    return readStream(memory.Get());
}

PartUri Package::uri(std::wstring_view partName) const {
    PartUri u;
    check(factory_->CreatePartUri(std::wstring(partName).c_str(), &u), "Ungültiger Teilname: " + narrow(partName));
    return u;
}

std::wstring Package::name(IOpcUri* uri) {
    Bstr path;
    check(uri->GetPath(path.put()), "Teilnamen lesen");
    return path.str();
}

bool Package::exists(IOpcPartUri* uri) const {
    BOOL found = FALSE;
    check(parts_->PartExists(uri, &found), "Teil suchen");
    return found != FALSE;
}

ComPtr<IOpcPart> Package::part(IOpcPartUri* uri) const {
    ComPtr<IOpcPart> p;
    check(parts_->GetPart(uri, &p), "Teil fehlt: " + narrow(name(uri)));
    return p;
}

ComPtr<IOpcPart> Package::createPart(std::wstring_view partName, std::wstring_view contentType) {
    ComPtr<IOpcPart> p;
    check(parts_->CreatePart(uri(partName).Get(), std::wstring(contentType).c_str(), OPC_COMPRESSION_NORMAL, &p),
          "Teil anlegen: " + narrow(partName));
    return p;
}

ComPtr<IStream> Package::stream(IOpcPart* part) const {
    ComPtr<IStream> s;
    check(part->GetContentStream(&s), "Teilinhalt öffnen");
    return s;
}

std::vector<std::uint8_t> Package::read(IOpcPart* part) const {
    return readStream(stream(part).Get());
}

void Package::write(IOpcPart* part, const std::vector<std::uint8_t>& bytes) const {
    writeStream(stream(part).Get(), bytes);
}

void Package::writeText(IOpcPart* part, std::wstring_view xml) const {
    const std::string utf8 = narrow(xml);
    write(part, std::vector<std::uint8_t>(utf8.begin(), utf8.end()));
}

std::wstring Package::uniqueName(std::wstring_view like) const {
    std::wstring s(like);
    const auto slash = s.find_last_of(L'/');
    const auto dot = s.find_last_of(L'.');
    const std::wstring ext = (dot != std::wstring::npos && (slash == std::wstring::npos || dot > slash)) ? s.substr(dot) : L"";
    std::wstring stem = s.substr(0, s.size() - ext.size());
    while (!stem.empty() && std::iswdigit(stem.back())) stem.pop_back();
    for (int n = 1;; ++n) {
        std::wstring candidate = stem + std::to_wstring(n) + ext;
        if (!exists(uri(candidate).Get())) return candidate;
    }
}

ComPtr<IOpcRelationshipSet> Package::relationshipSet(IOpcPart* source) const {
    ComPtr<IOpcRelationshipSet> set;
    if (source)
        check(source->GetRelationshipSet(&set), "Beziehungen lesen");
    else
        check(package_->GetRelationshipSet(&set), "Paketbeziehungen lesen");
    return set;
}

ComPtr<IOpcUri> Package::sourceUri(IOpcPart* source) const {
    if (source) {
        PartUri partUri;
        check(source->GetName(&partUri), "Teilnamen lesen");
        return ComPtr<IOpcUri>(partUri);
    }
    ComPtr<IOpcUri> root;
    check(factory_->CreatePackageRootUri(&root), "Paketwurzel bestimmen");
    return root;
}

std::vector<Relationship> Package::relationships(IOpcPart* source) const {
    auto set = relationshipSet(source);
    auto base = sourceUri(source);
    ComPtr<IOpcRelationshipEnumerator> e;
    check(set->GetEnumerator(&e), "Beziehungen aufzählen");
    std::vector<Relationship> out;
    BOOL more = FALSE;
    while (SUCCEEDED(e->MoveNext(&more)) && more) {
        ComPtr<IOpcRelationship> r;
        check(e->GetCurrent(&r), "Beziehung lesen");
        Relationship rel;
        CoStr id, type;
        check(r->GetId(id.put()), "Beziehungs-ID lesen");
        check(r->GetRelationshipType(type.put()), "Beziehungstyp lesen");
        rel.id = id.str();
        rel.type = type.str();
        OPC_URI_TARGET_MODE mode{};
        check(r->GetTargetMode(&mode), "Beziehungsziel lesen");
        rel.external = mode == OPC_URI_TARGET_MODE_EXTERNAL;
        check(r->GetTargetUri(&rel.targetUri), "Beziehungsziel lesen");
        if (!rel.external && FAILED(base->CombinePartUri(rel.targetUri.Get(), &rel.targetPart))) rel.targetPart.Reset();
        out.push_back(std::move(rel));
    }
    return out;
}

std::wstring Package::addRelationship(IOpcPart* source, std::wstring_view type, IOpcPartUri* target, std::wstring_view id) {
    ComPtr<IUri> relative;
    check(sourceUri(source)->GetRelativeUri(target, &relative), "Relatives Ziel bestimmen");
    const std::wstring idText(id);
    const std::wstring typeText(type);
    ComPtr<IOpcRelationship> r;
    check(relationshipSet(source)->CreateRelationship(idText.empty() ? nullptr : idText.c_str(), typeText.c_str(),
                                                      relative.Get(), OPC_URI_TARGET_MODE_INTERNAL, &r),
          "Beziehung anlegen");
    CoStr newId;
    check(r->GetId(newId.put()), "Beziehungs-ID lesen");
    return newId.str();
}

void Package::deleteRelationship(IOpcPart* source, std::wstring_view id) {
    relationshipSet(source)->DeleteRelationship(std::wstring(id).c_str());
}

PartUri Package::copyPartDeep(IOpcPartUri* source, CopyMemo& memo) {
    const std::wstring key = toLower(name(source));
    if (auto it = memo.find(key); it != memo.end()) return it->second;

    auto from = part(source);
    CoStr contentType;
    check(from->GetContentType(contentType.put()), "Inhaltstyp lesen");
    OPC_COMPRESSION_OPTIONS compression{};
    check(from->GetCompressionOptions(&compression), "Kompression lesen");

    const std::wstring newName = uniqueName(name(source));
    PartUri targetUri = uri(newName);
    ComPtr<IOpcPart> to;
    check(parts_->CreatePart(targetUri.Get(), contentType.str().c_str(), compression, &to), "Teil anlegen: " + narrow(newName));
    memo[key] = targetUri;
    write(to.Get(), read(from.Get()));

    auto targetSet = relationshipSet(to.Get());
    for (auto& rel : relationships(from.Get())) {
        const CopyRule rule = rel.external ? CopyRule::Share : ruleFor(rel.type);
        if (rule == CopyRule::Drop) continue;
        if (rule == CopyRule::Share || !rel.targetPart || !exists(rel.targetPart.Get())) {
            // Gleiches Ziel wie im Original (relativ zum selben Ordner, daher unverändert gültig).
            ComPtr<IOpcRelationship> created;
            check(targetSet->CreateRelationship(rel.id.c_str(), rel.type.c_str(), rel.targetUri.Get(),
                                                rel.external ? OPC_URI_TARGET_MODE_EXTERNAL : OPC_URI_TARGET_MODE_INTERNAL,
                                                &created),
                  "Beziehung kopieren");
            continue;
        }
        PartUri copy = copyPartDeep(rel.targetPart.Get(), memo);
        addRelationship(to.Get(), rel.type, copy.Get(), rel.id);
    }
    return targetUri;
}

size_t Package::collectGarbage() {
    std::set<std::wstring> reachable;
    std::vector<PartUri> pending;
    for (auto& rel : relationships(nullptr))
        if (rel.targetPart) pending.push_back(rel.targetPart);
    while (!pending.empty()) {
        PartUri current = pending.back();
        pending.pop_back();
        if (!reachable.insert(toLower(name(current.Get()))).second || !exists(current.Get())) continue;
        for (auto& rel : relationships(part(current.Get()).Get()))
            if (rel.targetPart) pending.push_back(rel.targetPart);
    }

    std::vector<PartUri> unreachable;
    ComPtr<IOpcPartEnumerator> e;
    check(parts_->GetEnumerator(&e), "Teile aufzählen");
    BOOL more = FALSE;
    while (SUCCEEDED(e->MoveNext(&more)) && more) {
        ComPtr<IOpcPart> p;
        check(e->GetCurrent(&p), "Teil lesen");
        PartUri partUri;
        check(p->GetName(&partUri), "Teilnamen lesen");
        if (!reachable.contains(toLower(name(partUri.Get())))) unreachable.push_back(partUri);
    }
    for (auto& u : unreachable) check(parts_->DeletePart(u.Get()), "Verwaisten Teil löschen");
    return unreachable.size();
}

}  // namespace rethink
