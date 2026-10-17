#include "dispatch.h"

#include <format>
#include <vector>

namespace rethink::disp {

Value Value::from(long v) {
    Value r;
    r.v_.vt = VT_I4;
    r.v_.lVal = v;
    return r;
}

Value Value::from(double v) {
    Value r;
    r.v_.vt = VT_R8;
    r.v_.dblVal = v;
    return r;
}

Value Value::from(bool v) {
    Value r;
    r.v_.vt = VT_BOOL;
    r.v_.boolVal = v ? VARIANT_TRUE : VARIANT_FALSE;
    return r;
}

Value Value::from(std::wstring_view v) {
    Value r;
    r.v_.vt = VT_BSTR;
    r.v_.bstrVal = SysAllocStringLen(v.data(), static_cast<UINT>(v.size()));
    return r;
}

Value Value::from(IDispatch* v) {
    Value r;
    r.v_.vt = VT_DISPATCH;
    r.v_.pdispVal = v;
    if (v) v->AddRef();
    return r;
}

ComPtr<IDispatch> Value::object() const {
    ComPtr<IDispatch> result;
    if (v_.vt == VT_DISPATCH && v_.pdispVal) {
        result = v_.pdispVal;
    } else if (v_.vt == VT_UNKNOWN && v_.punkVal) {
        v_.punkVal->QueryInterface(IID_PPV_ARGS(&result));
    }
    if (!result) throw Error("PowerPoint hat kein Objekt geliefert");
    return result;
}

std::wstring Value::str() const {
    VARIANT copy;
    VariantInit(&copy);
    if (FAILED(VariantChangeType(&copy, &v_, VARIANT_ALPHABOOL, VT_BSTR))) return {};
    std::wstring s(copy.bstrVal, SysStringLen(copy.bstrVal));
    VariantClear(&copy);
    return s;
}

long Value::toLong() const {
    VARIANT copy;
    VariantInit(&copy);
    check(VariantChangeType(&copy, &v_, 0, VT_I4), "Wert in Zahl umwandeln");
    return copy.lVal;
}

double Value::toDouble() const {
    VARIANT copy;
    VariantInit(&copy);
    check(VariantChangeType(&copy, &v_, 0, VT_R8), "Wert in Zahl umwandeln");
    return copy.dblVal;
}

Value invoke(IDispatch* target, std::wstring_view name, WORD flags, Value* args, size_t count) {
    if (!target) throw Error(std::format("{}: kein Objekt", narrow(name)));
    std::wstring member(name);
    LPOLESTR names[] = {member.data()};
    DISPID id = DISPID_UNKNOWN;
    check(target->GetIDsOfNames(IID_NULL, names, 1, LOCALE_USER_DEFAULT, &id), "PowerPoint-Element " + narrow(name));

    // IDispatch erwartet die Argumente in umgekehrter Reihenfolge.
    std::vector<VARIANT> reversed(count);
    for (size_t i = 0; i < count; ++i) reversed[count - 1 - i] = args[i].raw();
    DISPID putId = DISPID_PROPERTYPUT;
    DISPPARAMS params{reversed.data(), nullptr, static_cast<UINT>(count), 0};
    if (flags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) {
        params.rgdispidNamedArgs = &putId;
        params.cNamedArgs = 1;
    }

    Value result;
    EXCEPINFO info{};
    UINT argError = 0;
    const HRESULT hr = target->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, flags, &params, result.put(), &info, &argError);
    if (hr == DISP_E_EXCEPTION) {
        if (info.pfnDeferredFillIn) info.pfnDeferredFillIn(&info);
        std::wstring description = info.bstrDescription ? info.bstrDescription : L"";
        const HRESULT code = info.scode ? info.scode : static_cast<HRESULT>(info.wCode);
        SysFreeString(info.bstrSource);
        SysFreeString(info.bstrDescription);
        SysFreeString(info.bstrHelpFile);
        while (!description.empty() && (description.back() == L'\n' || description.back() == L'\r' || description.back() == L' '))
            description.pop_back();
        if (description.empty()) throwHr(code, "PowerPoint: " + narrow(name));
        throw Error(std::format("PowerPoint ({}): {}", narrow(name), narrow(description)));
    }
    check(hr, "PowerPoint: " + narrow(name));
    return result;
}

}  // namespace rethink::disp
