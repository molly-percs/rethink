#include "fill_text.h"

#include <format>

namespace rethink {
namespace {

std::vector<std::wstring> splitLines(const std::wstring& text) {
    std::vector<std::wstring> lines;
    size_t start = 0;
    for (;;) {
        const size_t end = text.find(L'\n', start);
        std::wstring line = text.substr(start, end == std::wstring::npos ? std::wstring::npos : end - start);
        if (!line.empty() && line.back() == L'\r') line.pop_back();
        lines.push_back(std::move(line));
        if (end == std::wstring::npos) return lines;
        start = end + 1;
    }
}

Node makeRun(const XmlDoc& doc, IXMLDOMNode* runProperties, const std::wstring& text) {
    Node run = doc.fragment(std::format(L"<a:r xmlns:a=\"{}\"><a:t>{}</a:t></a:r>", kNsA, xmlEscapeText(text)));
    if (runProperties) {
        Node t = doc.one(run.Get(), L"a:t");
        insertBefore(run.Get(), cloneDeep(runProperties).Get(), t.Get());
    }
    return run;
}

// Ersetzt den Inhalt eines Textkörpers durch `text` und übernimmt die Formatierung des ersten Laufs.
void setBodyText(const XmlDoc& doc, IXMLDOMNode* body, const std::wstring& text) {
    auto paragraphs = doc.all(body, L"a:p");
    if (paragraphs.empty()) {
        Node p = doc.fragment(std::format(L"<a:p xmlns:a=\"{}\"/>", kNsA));
        appendChild(body, p.Get());
        paragraphs.push_back(p);
    }
    Node first = paragraphs.front();
    Node runProperties = doc.one(first.Get(), L"a:r/a:rPr");
    if (!runProperties) runProperties = doc.one(first.Get(), L"a:fld/a:rPr");
    if (runProperties) runProperties = cloneDeep(runProperties.Get());

    for (auto& child : childElements(first.Get())) {
        if (isElement(child.Get(), kNsA, L"pPr") || isElement(child.Get(), kNsA, L"endParaRPr")) continue;
        removeNode(child.Get());
    }
    for (size_t i = 1; i < paragraphs.size(); ++i) removeNode(paragraphs[i].Get());

    const auto lines = splitLines(text);
    Node template_ = cloneDeep(first.Get());  // leerer Absatz mit pPr/endParaRPr
    Node previous = first;
    for (size_t i = 0; i < lines.size(); ++i) {
        Node paragraph = i == 0 ? first : cloneDeep(template_.Get());
        Node end = doc.one(paragraph.Get(), L"a:endParaRPr");
        insertBefore(paragraph.Get(), makeRun(doc, runProperties.Get(), lines[i]).Get(), end.Get());
        if (i > 0) {
            insertAfter(previous.Get(), paragraph.Get());
            previous = paragraph;
        }
    }
}

}  // namespace

bool matchesName(const XmlDoc& doc, IXMLDOMNode* cNvPr, const std::wstring& name) {
    if (!cNvPr) return false;
    return doc.attr(cNvPr, L"name") == name || doc.attr(cNvPr, L"descr") == L"rethink:" + name;
}

int fillTextInSlide(const XmlDoc& slide, const DataItem& item, std::vector<std::string>& warnings) {
    const bool single = item.table.size() == 1 && item.table.front().size() == 1;
    const std::wstring text = (!item.table.empty() && !item.table.front().empty()) ? cellText(item.table.front().front()) : L"";
    int filled = 0;

    // 1) Felder aus think-cell-Vorlagen: Der Name steht im eingebetteten XML des type-Attributs.
    const std::wstring marker = L"<m_strName>" + xmlEscapeText(item.name) + L"</m_strName>";
    for (auto& field : slide.all(nullptr, L"//a:fld[starts-with(@type,'thinkcell')]")) {
        if (slide.attr(field.Get(), L"type").find(marker) == std::wstring::npos) continue;
        Node runProperties = slide.one(field.Get(), L"a:rPr");
        replaceNode(field.Get(), makeRun(slide, runProperties.Get(), text).Get());
        ++filled;
    }

    // 2) Benannte Formen.
    for (auto& shape : slide.all(nullptr, L"//p:sp")) {
        if (!matchesName(slide, slide.one(shape.Get(), L"p:nvSpPr/p:cNvPr").Get(), item.name)) continue;
        Node body = slide.one(shape.Get(), L"p:txBody");
        if (!body) {
            warnings.push_back(std::format("Form '{}' hat keinen Text und wurde übersprungen", narrow(item.name)));
            continue;
        }
        setBodyText(slide, body.Get(), text);
        ++filled;
    }

    if (filled > 0 && !single)
        warnings.push_back(std::format("Textfeld '{}' erwartet genau eine Zelle; die erste wurde verwendet", narrow(item.name)));
    return filled;
}

}  // namespace rethink
