#pragma once

#include "util.h"

#include <msxml6.h>

#include <initializer_list>
#include <string>
#include <string_view>
#include <vector>

namespace rethink {

inline constexpr wchar_t kNsA[] = L"http://schemas.openxmlformats.org/drawingml/2006/main";
inline constexpr wchar_t kNsP[] = L"http://schemas.openxmlformats.org/presentationml/2006/main";
inline constexpr wchar_t kNsR[] = L"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
inline constexpr wchar_t kNsC[] = L"http://schemas.openxmlformats.org/drawingml/2006/chart";

using Node = ComPtr<IXMLDOMNode>;

// DOM eines OOXML-Teils (MSXML 6). XPath-Präfixe: a, p, r, c, mc.
class XmlDoc {
public:
    static XmlDoc load(IStream* stream);
    void save(IStream* stream) const;

    Node one(IXMLDOMNode* context, std::wstring_view xpath) const;
    std::vector<Node> all(IXMLDOMNode* context, std::wstring_view xpath) const;
    // Parst ein Fragment mit eigenen xmlns-Deklarationen und importiert es in dieses Dokument.
    Node fragment(std::wstring_view xml) const;
    // Attributwert oder leer; Präfixe wie "r:id" werden über XPath aufgelöst.
    std::wstring attr(IXMLDOMNode* node, std::wstring_view qualifiedName) const;

private:
    ComPtr<IXMLDOMDocument3> doc_;
};

std::wstring nodeText(IXMLDOMNode* node);
void setNodeText(IXMLDOMNode* node, std::wstring_view text);
void setAttr(IXMLDOMNode* node, std::wstring_view name, std::wstring_view value);
std::wstring localName(IXMLDOMNode* node);
bool isElement(IXMLDOMNode* node, std::wstring_view ns, std::wstring_view local);
Node cloneDeep(IXMLDOMNode* node);
std::vector<Node> childElements(IXMLDOMNode* parent);
// Erstes Kind-Element aus `ns`, dessen lokaler Name in `names` steht.
Node firstChildOf(IXMLDOMNode* parent, std::wstring_view ns, std::initializer_list<std::wstring_view> names);
Node nextSibling(IXMLDOMNode* node);
void appendChild(IXMLDOMNode* parent, IXMLDOMNode* child);
// reference == nullptr hängt an.
void insertBefore(IXMLDOMNode* parent, IXMLDOMNode* child, IXMLDOMNode* reference);
void insertAfter(IXMLDOMNode* node, IXMLDOMNode* newSibling);
void removeNode(IXMLDOMNode* node);
void replaceNode(IXMLDOMNode* oldNode, IXMLDOMNode* newNode);

}  // namespace rethink
