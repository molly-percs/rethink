#include "xmldoc.h"

#include <format>

namespace rethink {
namespace {

constexpr wchar_t kSelectionNamespaces[] =
    L"xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' "
    L"xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main' "
    L"xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' "
    L"xmlns:c='http://schemas.openxmlformats.org/drawingml/2006/chart' "
    L"xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006' "
    L"xmlns:s='http://schemas.openxmlformats.org/spreadsheetml/2006/main'";

void setProperty(IXMLDOMDocument3* doc, std::wstring_view name, std::wstring_view value) {
    Bstr n(name);
    Variant v;
    Variant::setString(v, value);
    check(doc->setProperty(n.get(), v.get()), "MSXML-Eigenschaft setzen");
}

ComPtr<IXMLDOMDocument3> newDocument() {
    ComPtr<IXMLDOMDocument3> doc;
    check(CoCreateInstance(__uuidof(DOMDocument60), nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&doc)), "MSXML 6 initialisieren");
    check(doc->put_async(VARIANT_FALSE), "MSXML konfigurieren");
    check(doc->put_validateOnParse(VARIANT_FALSE), "MSXML konfigurieren");
    check(doc->put_resolveExternals(VARIANT_FALSE), "MSXML konfigurieren");
    check(doc->put_preserveWhiteSpace(VARIANT_TRUE), "MSXML konfigurieren");
    setProperty(doc.Get(), L"SelectionLanguage", L"XPath");
    setProperty(doc.Get(), L"SelectionNamespaces", kSelectionNamespaces);
    return doc;
}

std::string parseErrorText(IXMLDOMDocument3* doc) {
    ComPtr<IXMLDOMParseError> error;
    if (FAILED(doc->get_parseError(&error)) || !error) return "unbekannter Fehler";
    Bstr reason;
    error->get_reason(reason.put());
    long line = 0;
    error->get_line(&line);
    return std::format("{} (Zeile {})", narrow(reason.str()), line);
}

}  // namespace

XmlDoc XmlDoc::load(IStream* stream) {
    LARGE_INTEGER zero{};
    check(stream->Seek(zero, STREAM_SEEK_SET, nullptr), "Stream positionieren");
    XmlDoc d;
    d.doc_ = newDocument();
    Variant source;
    Variant::setUnknown(source, stream);
    VARIANT_BOOL ok = VARIANT_FALSE;
    const HRESULT hr = d.doc_->load(source.get(), &ok);
    if (FAILED(hr) || ok != VARIANT_TRUE) throw Error("XML konnte nicht gelesen werden: " + parseErrorText(d.doc_.Get()));
    return d;
}

void XmlDoc::save(IStream* stream) const {
    ULARGE_INTEGER zeroSize{};
    check(stream->SetSize(zeroSize), "Stream leeren");
    LARGE_INTEGER zero{};
    check(stream->Seek(zero, STREAM_SEEK_SET, nullptr), "Stream positionieren");
    Variant target;
    Variant::setUnknown(target, stream);
    check(doc_->save(target.get()), "XML schreiben");
}

Node XmlDoc::one(IXMLDOMNode* context, std::wstring_view xpath) const {
    IXMLDOMNode* base = context ? context : static_cast<IXMLDOMNode*>(doc_.Get());
    Bstr query(xpath);
    Node node;
    const HRESULT hr = base->selectSingleNode(query.get(), &node);
    if (FAILED(hr)) throwHr(hr, "XPath " + narrow(xpath));
    return node;
}

std::vector<Node> XmlDoc::all(IXMLDOMNode* context, std::wstring_view xpath) const {
    IXMLDOMNode* base = context ? context : static_cast<IXMLDOMNode*>(doc_.Get());
    Bstr query(xpath);
    ComPtr<IXMLDOMNodeList> list;
    const HRESULT hr = base->selectNodes(query.get(), &list);
    if (FAILED(hr)) throwHr(hr, "XPath " + narrow(xpath));
    long length = 0;
    check(list->get_length(&length), "XPath-Ergebnis lesen");
    std::vector<Node> out;
    out.reserve(static_cast<size_t>(length));
    for (long i = 0; i < length; ++i) {
        Node item;
        check(list->get_item(i, &item), "XPath-Ergebnis lesen");
        out.push_back(item);
    }
    return out;
}

Node XmlDoc::fragment(std::wstring_view xml) const {
    auto temp = newDocument();
    Bstr source(xml);
    VARIANT_BOOL ok = VARIANT_FALSE;
    const HRESULT hr = temp->loadXML(source.get(), &ok);
    if (FAILED(hr) || ok != VARIANT_TRUE) throw Error("Internes XML-Fragment ungültig: " + parseErrorText(temp.Get()));
    ComPtr<IXMLDOMElement> root;
    check(temp->get_documentElement(&root), "XML-Fragment lesen");
    Node imported;
    check(doc_->importNode(root.Get(), VARIANT_TRUE, &imported), "XML-Fragment importieren");
    return imported;
}

std::wstring XmlDoc::attr(IXMLDOMNode* node, std::wstring_view qualifiedName) const {
    if (!node) return {};
    auto a = one(node, std::wstring(L"@") + std::wstring(qualifiedName));
    return a ? nodeText(a.Get()) : std::wstring();
}

std::wstring nodeText(IXMLDOMNode* node) {
    Bstr text;
    check(node->get_text(text.put()), "XML-Text lesen");
    return text.str();
}

void setNodeText(IXMLDOMNode* node, std::wstring_view text) {
    Bstr value(text);
    check(node->put_text(value.get()), "XML-Text setzen");
}

void setAttr(IXMLDOMNode* node, std::wstring_view name, std::wstring_view value) {
    ComPtr<IXMLDOMElement> element;
    check(node->QueryInterface(IID_PPV_ARGS(&element)), "XML-Element erwartet");
    Bstr n(name);
    Variant v;
    Variant::setString(v, value);
    check(element->setAttribute(n.get(), v.get()), "XML-Attribut setzen");
}

std::wstring localName(IXMLDOMNode* node) {
    Bstr name;
    check(node->get_baseName(name.put()), "XML-Namen lesen");
    return name.str();
}

bool isElement(IXMLDOMNode* node, std::wstring_view ns, std::wstring_view local) {
    DOMNodeType type{};
    if (FAILED(node->get_nodeType(&type)) || type != NODE_ELEMENT) return false;
    Bstr uri;
    node->get_namespaceURI(uri.put());
    return uri.str() == ns && localName(node) == local;
}

Node cloneDeep(IXMLDOMNode* node) {
    Node clone;
    check(node->cloneNode(VARIANT_TRUE, &clone), "XML-Knoten kopieren");
    return clone;
}

std::vector<Node> childElements(IXMLDOMNode* parent) {
    std::vector<Node> out;
    Node child;
    parent->get_firstChild(&child);
    while (child) {
        DOMNodeType type{};
        child->get_nodeType(&type);
        if (type == NODE_ELEMENT) out.push_back(child);
        Node next;
        child->get_nextSibling(&next);
        child = next;
    }
    return out;
}

Node firstChildOf(IXMLDOMNode* parent, std::wstring_view ns, std::initializer_list<std::wstring_view> names) {
    for (auto& child : childElements(parent)) {
        for (auto name : names) {
            if (isElement(child.Get(), ns, name)) return child;
        }
    }
    return nullptr;
}

Node nextSibling(IXMLDOMNode* node) {
    Node next;
    node->get_nextSibling(&next);
    return next;
}

void appendChild(IXMLDOMNode* parent, IXMLDOMNode* child) {
    Node added;
    check(parent->appendChild(child, &added), "XML-Knoten anhängen");
}

void insertBefore(IXMLDOMNode* parent, IXMLDOMNode* child, IXMLDOMNode* reference) {
    if (!reference) {
        appendChild(parent, child);
        return;
    }
    Variant ref;
    Variant::setDispatch(ref, reference);
    Node added;
    check(parent->insertBefore(child, ref.get(), &added), "XML-Knoten einfügen");
}

void insertAfter(IXMLDOMNode* node, IXMLDOMNode* newSibling) {
    Node parent;
    check(node->get_parentNode(&parent), "XML-Elternknoten lesen");
    insertBefore(parent.Get(), newSibling, nextSibling(node).Get());
}

void removeNode(IXMLDOMNode* node) {
    Node parent;
    node->get_parentNode(&parent);
    if (!parent) return;
    Node removed;
    check(parent->removeChild(node, &removed), "XML-Knoten entfernen");
}

void replaceNode(IXMLDOMNode* oldNode, IXMLDOMNode* newNode) {
    Node parent;
    check(oldNode->get_parentNode(&parent), "XML-Elternknoten lesen");
    Node replaced;
    check(parent->replaceChild(newNode, oldNode, &replaced), "XML-Knoten ersetzen");
}

}  // namespace rethink
