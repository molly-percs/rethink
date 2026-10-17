#pragma once

// Spätes Binden an das PowerPoint-Objektmodell über IDispatch (ohne Typbibliothek).

#include "util.h"

#include <array>
#include <string>
#include <string_view>

namespace rethink::disp {

// VARIANT mit Besitz; verschiebbar, nicht kopierbar.
class Value {
public:
    Value() { VariantInit(&v_); }
    ~Value() { VariantClear(&v_); }
    Value(Value&& other) noexcept : v_(other.v_) { other.v_.vt = VT_EMPTY; }
    Value& operator=(Value&& other) noexcept {
        if (this != &other) {
            VariantClear(&v_);
            v_ = other.v_;
            other.v_.vt = VT_EMPTY;
        }
        return *this;
    }
    Value(const Value&) = delete;
    Value& operator=(const Value&) = delete;

    static Value from(int v) { return from(static_cast<long>(v)); }
    static Value from(long v);
    static Value from(double v);
    static Value from(bool v);
    static Value from(std::wstring_view v);
    static Value from(const wchar_t* v) { return from(std::wstring_view(v)); }
    static Value from(const std::wstring& v) { return from(std::wstring_view(v)); }
    static Value from(IDispatch* v);

    ComPtr<IDispatch> object() const;  // wirft, wenn kein Objekt geliefert wurde
    std::wstring str() const;
    long toLong() const;
    double toDouble() const;

    VARIANT* put() {
        VariantClear(&v_);
        return &v_;
    }
    VARIANT& raw() { return v_; }

private:
    VARIANT v_;
};

// Ruft eine Methode oder Eigenschaft auf. args: in normaler Reihenfolge (erstes Argument zuerst).
Value invoke(IDispatch* target, std::wstring_view name, WORD flags, Value* args, size_t count);

template <class... A>
Value call(IDispatch* target, std::wstring_view name, A&&... args) {
    std::array<Value, sizeof...(A)> list{Value::from(std::forward<A>(args))...};
    return invoke(target, name, DISPATCH_METHOD | DISPATCH_PROPERTYGET, list.data(), list.size());
}

template <class... A>
ComPtr<IDispatch> object(IDispatch* target, std::wstring_view name, A&&... args) {
    return call(target, name, std::forward<A>(args)...).object();
}

template <class V>
void put(IDispatch* target, std::wstring_view name, V&& value) {
    Value v = Value::from(std::forward<V>(value));
    invoke(target, name, DISPATCH_PROPERTYPUT, &v, 1);
}

}  // namespace rethink::disp
