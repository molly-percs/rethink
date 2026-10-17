#include "render_element.h"

#include "engine/chart.h"
#include "fill_text.h"

#include <algorithm>
#include <cstdlib>
#include <format>
#include <stdexcept>

namespace rethink {
namespace {

using namespace engine;

constexpr wchar_t kRelTags[] = L"http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags";
constexpr wchar_t kTagsContentType[] = L"application/vnd.openxmlformats-officedocument.presentationml.tags+xml";

std::wstring colorXml(const Color& c) {
    switch (c.kind) {
        case Color::Kind::Scheme: return std::format(L"<a:solidFill><a:schemeClr val=\"{}\"/></a:solidFill>", c.scheme);
        case Color::Kind::Rgb: return std::format(L"<a:solidFill><a:srgbClr val=\"{:06X}\"/></a:solidFill>", c.rgb);
        case Color::Kind::None: break;
    }
    return L"<a:noFill/>";
}

std::wstring xfrmXml(Emu x, Emu y, Emu w, Emu h, bool flipH = false, bool flipV = false) {
    return std::format(L"<a:xfrm{}{}><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/></a:xfrm>", flipH ? L" flipH=\"1\"" : L"",
                       flipV ? L" flipV=\"1\"" : L"", x, y, w, h);
}

std::wstring primXml(const Prim& p, int id) {
    const std::wstring name = xmlEscape(std::wstring(p.role.begin(), p.role.end()));
    switch (p.kind) {
        case Prim::Kind::Rect:
            return std::format(
                L"<p:sp><p:nvSpPr><p:cNvPr id=\"{}\" name=\"{}\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>{}"
                L"<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>{}<a:ln><a:noFill/></a:ln></p:spPr></p:sp>",
                id, name, xfrmXml(p.x, p.y, p.w, p.h), colorXml(p.fill));
        case Prim::Kind::Line: {
            const bool flipH = p.w < 0, flipV = p.h < 0;
            const Emu x = flipH ? p.x + p.w : p.x, y = flipV ? p.y + p.h : p.y;
            return std::format(
                L"<p:cxnSp><p:nvCxnSpPr><p:cNvPr id=\"{}\" name=\"{}\"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>{}"
                L"<a:prstGeom prst=\"line\"><a:avLst/></a:prstGeom><a:ln w=\"{}\">{}{}</a:ln></p:spPr></p:cxnSp>",
                id, name, xfrmXml(x, y, std::abs(p.w), std::abs(p.h), flipH, flipV), static_cast<Emu>(p.strokePt * kEmuPerPt),
                colorXml(p.stroke), p.dashed ? L"<a:prstDash val=\"dash\"/>" : L"");
        }
        case Prim::Kind::Text: {
            const wchar_t* align = p.align == Align::Left ? L"l" : p.align == Align::Right ? L"r" : L"ctr";
            return std::format(
                L"<p:sp><p:nvSpPr><p:cNvPr id=\"{}\" name=\"{}\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr><p:spPr>{}"
                L"<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>"
                L"<p:txBody><a:bodyPr wrap=\"none\" lIns=\"0\" tIns=\"0\" rIns=\"0\" bIns=\"0\" anchor=\"ctr\"><a:noAutofit/></a:bodyPr>"
                L"<a:lstStyle/><a:p><a:pPr algn=\"{}\"/><a:r><a:rPr lang=\"de-DE\" sz=\"{}\" dirty=\"0\">{}<a:latin typeface=\"+mn-lt\"/>"
                L"</a:rPr><a:t>{}</a:t></a:r></a:p></p:txBody></p:sp>",
                id, name, xfrmXml(p.x, p.y, p.w, p.h), align, static_cast<int>(p.fontPt * 100), colorXml(p.textColor),
                xmlEscapeText(p.text));
        }
    }
    return {};
}

std::wstring groupXml(const Drawing& d, const std::wstring& name, const std::wstring& tagRelId, int& nextId) {
    std::wstring children;
    const int groupId = nextId++;
    for (auto& p : d.prims) children += primXml(p, nextId++);
    const Box& b = d.bounds;
    return std::format(
        L"<p:grpSp xmlns:p=\"{}\" xmlns:a=\"{}\" xmlns:r=\"{}\"><p:nvGrpSpPr><p:cNvPr id=\"{}\" name=\"{}\"/><p:cNvGrpSpPr/>"
        L"<p:nvPr><p:custDataLst><p:tags r:id=\"{}\"/></p:custDataLst></p:nvPr></p:nvGrpSpPr>"
        L"<p:grpSpPr><a:xfrm><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/><a:chOff x=\"{}\" y=\"{}\"/><a:chExt cx=\"{}\" cy=\"{}\"/></a:xfrm>"
        L"</p:grpSpPr>{}</p:grpSp>",
        kNsP, kNsA, kNsR, groupId, xmlEscape(name), tagRelId, b.x, b.y, b.w, b.h, b.x, b.y, b.w, b.h, children);
}

std::wstring tagXml(const std::wstring& model) {
    return std::format(
        L"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:tagLst xmlns:p=\"{}\"><p:tag name=\"{}\" val=\"{}\"/></p:tagLst>",
        kNsP, kModelTag, xmlEscape(model));
}

int maxShapeId(const XmlDoc& slide) {
    int maxId = 1;
    for (auto& n : slide.all(nullptr, L"//p:cNvPr")) {
        const std::wstring id = slide.attr(n.Get(), L"id");
        if (!id.empty()) maxId = std::max(maxId, std::stoi(id));
    }
    return maxId;
}

Box boxOf(const XmlDoc& slide, IXMLDOMNode* xfrm) {
    if (!xfrm) throw Error("Platzhalter hat keine Position");
    Node off = slide.one(xfrm, L"a:off"), ext = slide.one(xfrm, L"a:ext");
    if (!off || !ext) throw Error("Platzhalter hat keine Position");
    return {std::stoll(slide.attr(off.Get(), L"x")), std::stoll(slide.attr(off.Get(), L"y")), std::stoll(slide.attr(ext.Get(), L"cx")),
            std::stoll(slide.attr(ext.Get(), L"cy"))};
}

PartUri targetOf(const std::vector<Relationship>& rels, const std::wstring& id) {
    for (auto& r : rels)
        if (r.id == id) return r.targetPart;
    return nullptr;
}

// Liest das RETHINK_MODEL einer Gruppe; leer, wenn die Gruppe kein rethink-Element ist.
std::wstring storedModel(Package& pkg, const XmlDoc& slide, const std::vector<Relationship>& rels, IXMLDOMNode* group,
                         std::wstring& tagRelId) {
    for (auto& tags : slide.all(group, L"p:nvGrpSpPr/p:nvPr/p:custDataLst/p:tags")) {
        const std::wstring id = slide.attr(tags.Get(), L"r:id");
        PartUri target = targetOf(rels, id);
        if (!target || !pkg.exists(target.Get())) continue;
        auto stream = pkg.stream(pkg.part(target.Get()).Get());
        XmlDoc tagDoc = XmlDoc::load(stream.Get());
        for (auto& tag : tagDoc.all(nullptr, L"/p:tagLst/p:tag")) {
            if (tagDoc.attr(tag.Get(), L"name") == kModelTag) {
                tagRelId = id;
                return tagDoc.attr(tag.Get(), L"val");
            }
        }
    }
    return {};
}

std::wstring localeDecimalSeparator() {
    wchar_t buffer[8] = {};
    if (GetLocaleInfoEx(LOCALE_NAME_USER_DEFAULT, LOCALE_SDECIMAL, buffer, 8) > 0) return buffer;
    return L",";
}

Drawing render(const DataItem& item, const std::wstring& chartJson, Box bounds) {
    ChartOptions options = parseChartOptions(chartJson);
    options.decimalSeparator = localeDecimalSeparator();
    try {
        return layoutChart(item, options, bounds);
    } catch (const std::invalid_argument& e) {
        throw Error(std::format("Diagramm '{}': {}", narrow(item.name), e.what()));
    }
}

}  // namespace

int renderChartElements(Package& pkg, IOpcPart* slidePart, const XmlDoc& slide, std::vector<Relationship>& slideRels,
                        const DataItem& item, std::vector<std::string>& warnings) {
    int rendered = 0;
    int nextId = maxShapeId(slide) + 1;

    // 1) Bestehende rethink-Elemente neu zeichnen.
    for (auto& group : slide.all(nullptr, L"//p:grpSp")) {
        if (!matchesName(slide, slide.one(group.Get(), L"p:nvGrpSpPr/p:cNvPr").Get(), item.name)) continue;
        std::wstring tagRelId;
        const std::wstring stored = storedModel(pkg, slide, slideRels, group.Get(), tagRelId);
        if (stored.empty()) continue;

        DataItem merged = item;
        merged.chartJson = mergeChartJson(parseDataItemJson(stored).chartJson, item.chartJson);
        Drawing d = render(merged, merged.chartJson, boxOf(slide, slide.one(group.Get(), L"p:grpSpPr/a:xfrm").Get()));
        for (auto& w : d.warnings) warnings.push_back(std::format("Diagramm '{}': {}", narrow(item.name), w));

        pkg.writeText(pkg.part(targetOf(slideRels, tagRelId).Get()).Get(), tagXml(serializeDataItem(merged)));
        replaceNode(group.Get(), slide.fragment(groupXml(d, item.name, tagRelId, nextId)).Get());
        ++rendered;
    }

    // 2) Platzhalterformen in neue rethink-Elemente umwandeln.
    if (!item.chartJson.empty()) {
        for (auto& shape : slide.all(nullptr, L"//p:sp")) {
            if (!matchesName(slide, slide.one(shape.Get(), L"p:nvSpPr/p:cNvPr").Get(), item.name)) continue;
            Drawing d = render(item, item.chartJson, boxOf(slide, slide.one(shape.Get(), L"p:spPr/a:xfrm").Get()));
            for (auto& w : d.warnings) warnings.push_back(std::format("Diagramm '{}': {}", narrow(item.name), w));

            const std::wstring tagName = pkg.uniqueName(L"/ppt/tags/tag1.xml");
            auto tagPart = pkg.createPart(tagName, kTagsContentType);
            pkg.writeText(tagPart.Get(), tagXml(serializeDataItem(item)));
            const std::wstring relId = pkg.addRelationship(slidePart, kRelTags, pkg.uri(tagName).Get());
            slideRels = pkg.relationships(slidePart);

            replaceNode(shape.Get(), slide.fragment(groupXml(d, item.name, relId, nextId)).Get());
            ++rendered;
        }
    }
    return rendered;
}

}  // namespace rethink
