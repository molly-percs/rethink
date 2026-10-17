#include "build.h"

#include "fill_chart.h"
#include "fill_text.h"
#include "opc.h"
#include "ppttc.h"
#include "render_element.h"
#include "xmldoc.h"

#include <urlmon.h>

#include <algorithm>
#include <format>
#include <set>

namespace rethink {
namespace {

constexpr wchar_t kRelOfficeDocument[] = L"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
constexpr wchar_t kRelSlide[] = L"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
// p14:creationId; muss pro Folie eindeutig sein und wird bei Kopien entfernt.
constexpr wchar_t kCreationIdExt[] = L"{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}";

bool isUrl(const std::wstring& ref) { return ref.starts_with(L"http://") || ref.starts_with(L"https://"); }

// Heruntergeladene Vorlagen werden am Ende gelöscht.
class TempFiles {
public:
    ~TempFiles() {
        for (auto& f : files_) {
            std::error_code ec;
            std::filesystem::remove(f, ec);
        }
    }
    std::filesystem::path add() {
        wchar_t dir[MAX_PATH + 1] = {};
        GetTempPathW(MAX_PATH, dir);
        auto path = std::filesystem::path(dir) / std::format(L"rethink-{}-{}.pptx", GetCurrentProcessId(), files_.size());
        files_.push_back(path);
        return path;
    }

private:
    std::vector<std::filesystem::path> files_;
};

std::filesystem::path resolveTemplate(const std::wstring& ref, const std::filesystem::path& baseDir, TempFiles& temps) {
    if (isUrl(ref)) {
        auto target = temps.add();
        const HRESULT hr = URLDownloadToFileW(nullptr, ref.c_str(), target.c_str(), 0, nullptr);
        if (FAILED(hr)) throwHr(hr, "Vorlage herunterladen: " + narrow(ref));
        return target;
    }
    std::filesystem::path p(ref);
    if (p.is_relative()) p = baseDir / p;
    if (!std::filesystem::exists(p)) throw Error("Vorlage nicht gefunden: " + narrow(p.wstring()));
    return std::filesystem::weakly_canonical(p);
}

std::wstring templateKey(const std::wstring& ref, const std::filesystem::path& baseDir) {
    if (isUrl(ref)) return ref;
    std::filesystem::path p(ref);
    if (p.is_relative()) p = baseDir / p;
    return toLower(std::filesystem::weakly_canonical(p).wstring());
}

struct SlideDoc {
    ComPtr<IOpcPart> part;
    ComPtr<IStream> stream;
    XmlDoc doc;
    std::vector<Relationship> rels;
};

PartUri relationshipTarget(const std::vector<Relationship>& rels, const std::wstring& id) {
    for (auto& r : rels)
        if (r.id == id) return r.targetPart;
    return nullptr;
}

std::string notFoundMessage(const DataItem& item, const std::vector<SlideDoc>& slides) {
    std::string message = std::format("Element '{}' wurde in der Vorlage nicht gefunden", narrow(item.name));
    for (auto& s : slides) {
        if (s.doc.one(nullptr, L"//p:oleObj[starts-with(@progId,'TCLayout.')]")) {
            message +=
                " (die Vorlage enthält think-cell-Elemente; deren Diagrammnamen liegen in einem proprietären Datenobjekt."
                " Benenne das Diagramm im Auswahlbereich um oder setze den Alternativtext auf 'rethink:" +
                narrow(item.name) + "')";
            break;
        }
    }
    return message;
}

}  // namespace

BuildResult build(const std::filesystem::path& input, const std::filesystem::path& output) {
    BuildResult result;
    const auto specs = parsePpttc(input);
    if (specs.empty()) throw Error("Die .ppttc-Datei enthält keine Folien");

    const auto baseDir = std::filesystem::absolute(input).parent_path();
    const std::wstring key = templateKey(specs.front().templateRef, baseDir);
    for (auto& spec : specs) {
        if (templateKey(spec.templateRef, baseDir) != key)
            throw Error(std::format("rethink Phase 1 unterstützt nur eine Vorlage pro .ppttc-Datei (gefunden: '{}' und '{}')",
                                    narrow(specs.front().templateRef), narrow(spec.templateRef)));
    }

    TempFiles temps;
    const auto templatePath = resolveTemplate(specs.front().templateRef, baseDir, temps);
    if (std::filesystem::exists(output) && std::filesystem::equivalent(output, templatePath))
        throw Error("Die Ausgabedatei darf nicht die Vorlage selbst sein");

    Package pkg = Package::open(templatePath);

    PartUri presentationUri;
    for (auto& r : pkg.relationships(nullptr))
        if (r.type == kRelOfficeDocument) presentationUri = r.targetPart;
    if (!presentationUri) throw Error("Die Vorlage ist keine PowerPoint-Präsentation");
    auto presentationPart = pkg.part(presentationUri.Get());
    auto presentationStream = pkg.stream(presentationPart.Get());
    XmlDoc presentation = XmlDoc::load(presentationStream.Get());
    const auto presentationRels = pkg.relationships(presentationPart.Get());

    Node slideList = presentation.one(nullptr, L"/p:presentation/p:sldIdLst");
    std::vector<PartUri> originals;
    unsigned long maxId = 255;
    if (slideList) {
        for (auto& entry : presentation.all(slideList.Get(), L"p:sldId")) {
            maxId = std::max(maxId, std::stoul(presentation.attr(entry.Get(), L"id")));
            if (auto target = relationshipTarget(presentationRels, presentation.attr(entry.Get(), L"r:id"))) originals.push_back(target);
        }
    }
    if (originals.empty()) throw Error("Die Vorlage enthält keine Folien");

    // Eintrag 0 nutzt die Originalfolien, jeder weitere Eintrag eine vollständige Kopie.
    std::vector<std::vector<PartUri>> slidesPerEntry(specs.size());
    slidesPerEntry.front() = originals;
    for (size_t e = 1; e < specs.size(); ++e) {
        for (auto& original : originals) {
            Package::CopyMemo memo;
            PartUri copy = pkg.copyPartDeep(original.Get(), memo);
            const std::wstring id = pkg.addRelationship(presentationPart.Get(), kRelSlide, copy.Get());
            Node entry = presentation.fragment(std::format(L"<p:sldId xmlns:p=\"{}\" xmlns:r=\"{}\" id=\"{}\" r:id=\"{}\"/>",
                                                           kNsP, kNsR, ++maxId, id));
            appendChild(slideList.Get(), entry.Get());
            slidesPerEntry[e].push_back(copy);
        }
    }
    presentation.save(presentationStream.Get());

    std::set<std::wstring> filledCharts;  // Choice/Fallback-Paare verweisen auf dasselbe Diagramm
    for (size_t e = 0; e < specs.size(); ++e) {
        std::vector<SlideDoc> slides;
        for (auto& u : slidesPerEntry[e]) {
            auto part = pkg.part(u.Get());
            auto stream = pkg.stream(part.Get());
            XmlDoc doc = XmlDoc::load(stream.Get());
            slides.push_back({part, stream, std::move(doc), pkg.relationships(part.Get())});
        }
        if (e > 0) {
            for (auto& s : slides)
                if (Node ext = s.doc.one(nullptr, std::format(L"/p:sld/p:extLst/p:ext[@uri='{}']", kCreationIdExt))) removeNode(ext.Get());
        }

        for (auto& item : specs[e].data) {
            int hits = 0;
            for (auto& s : slides) {
                // Gezeichnete rethink-Diagramme (bestehende Elemente oder Platzhalter mit "chart"-Objekt)
                hits += renderChartElements(pkg, s.part.Get(), s.doc, s.rels, item, result.warnings);
                if (item.chartJson.empty()) hits += fillTextInSlide(s.doc, item, result.warnings);
                for (auto& frame : s.doc.all(nullptr, L"//p:graphicFrame")) {
                    if (!matchesName(s.doc, s.doc.one(frame.Get(), L"p:nvGraphicFramePr/p:cNvPr").Get(), item.name)) continue;
                    Node chartRef = s.doc.one(frame.Get(), L".//c:chart");
                    if (!chartRef) {
                        result.warnings.push_back(std::format("Element '{}' ist kein Diagramm und wurde übersprungen", narrow(item.name)));
                        continue;
                    }
                    PartUri chartUri = relationshipTarget(s.rels, s.doc.attr(chartRef.Get(), L"r:id"));
                    if (!chartUri) {
                        result.warnings.push_back(std::format("Diagramm '{}': Diagrammteil fehlt", narrow(item.name)));
                        continue;
                    }
                    if (filledCharts.insert(toLower(Package::name(chartUri.Get()))).second)
                        fillChart(pkg, pkg.part(chartUri.Get()).Get(), item, result.warnings);
                    ++hits;
                }
            }
            if (hits == 0) result.warnings.push_back(notFoundMessage(item, slides));
            result.filled += hits;
        }
        for (auto& s : slides) s.doc.save(s.stream.Get());
        result.slides += static_cast<int>(slides.size());
    }

    result.removedParts = pkg.collectGarbage();
    const auto outputDir = std::filesystem::absolute(output).parent_path();
    if (!outputDir.empty()) std::filesystem::create_directories(outputDir);
    if (std::filesystem::exists(output)) std::filesystem::remove(output);
    pkg.save(output);
    return result;
}

}  // namespace rethink
