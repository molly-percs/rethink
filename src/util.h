#pragma once

#include <windows.h>
#include <oleauto.h>
#include <wrl/client.h>

#include <cstdint>
#include <stdexcept>
#include <string>
#include <string_view>
#include <vector>

namespace rethink {

using Microsoft::WRL::ComPtr;

// Fehler, die dem Nutzer als Klartext gemeldet werden.
class Error : public std::runtime_error {
public:
    explicit Error(const std::string& message) : std::runtime_error(message) {}
};

std::wstring widen(std::string_view utf8);
std::string narrow(std::wstring_view wide);

[[noreturn]] void throwHr(HRESULT hr, std::string_view what);
inline void check(HRESULT hr, std::string_view what) {
    if (FAILED(hr)) throwHr(hr, what);
}

std::wstring toLower(std::wstring_view s);
// Escaping für Attributwerte und Textinhalt.
std::wstring xmlEscape(std::wstring_view s);
// Escaping nur für Textinhalt (&, <, >).
std::wstring xmlEscapeText(std::wstring_view s);
// Kürzeste Darstellung, die beim Einlesen denselben Wert ergibt; immer mit Punkt.
std::wstring formatNumber(double value);

std::vector<std::uint8_t> readStream(IStream* stream);
void writeStream(IStream* stream, const std::vector<std::uint8_t>& bytes);

class Bstr {
public:
    Bstr() = default;
    explicit Bstr(std::wstring_view s) : b_(SysAllocStringLen(s.data(), static_cast<UINT>(s.size()))) {}
    ~Bstr() { SysFreeString(b_); }
    Bstr(const Bstr&) = delete;
    Bstr& operator=(const Bstr&) = delete;

    BSTR* put() {
        SysFreeString(b_);
        b_ = nullptr;
        return &b_;
    }
    BSTR get() const { return b_; }
    std::wstring str() const { return b_ ? std::wstring(b_, SysStringLen(b_)) : std::wstring(); }

private:
    BSTR b_ = nullptr;
};

// Von OPC per CoTaskMemAlloc gelieferte Zeichenketten.
class CoStr {
public:
    CoStr() = default;
    ~CoStr() { CoTaskMemFree(p_); }
    CoStr(const CoStr&) = delete;
    CoStr& operator=(const CoStr&) = delete;

    LPWSTR* put() {
        CoTaskMemFree(p_);
        p_ = nullptr;
        return &p_;
    }
    std::wstring str() const { return p_ ? std::wstring(p_) : std::wstring(); }

private:
    LPWSTR p_ = nullptr;
};

class Variant {
public:
    Variant() { VariantInit(&v_); }
    ~Variant() { VariantClear(&v_); }
    Variant(const Variant&) = delete;
    Variant& operator=(const Variant&) = delete;

    static void setString(Variant& var, std::wstring_view s) {
        VariantClear(&var.v_);
        var.v_.vt = VT_BSTR;
        var.v_.bstrVal = SysAllocStringLen(s.data(), static_cast<UINT>(s.size()));
    }
    static void setUnknown(Variant& var, IUnknown* p) {
        VariantClear(&var.v_);
        var.v_.vt = VT_UNKNOWN;
        var.v_.punkVal = p;
        p->AddRef();
    }
    static void setDispatch(Variant& var, IDispatch* p) {
        VariantClear(&var.v_);
        var.v_.vt = VT_DISPATCH;
        var.v_.pdispVal = p;
        p->AddRef();
    }
    static void setInt(Variant& var, long value) {
        VariantClear(&var.v_);
        var.v_.vt = VT_I4;
        var.v_.lVal = value;
    }
    const VARIANT& get() const { return v_; }

private:
    VARIANT v_;
};

}  // namespace rethink
