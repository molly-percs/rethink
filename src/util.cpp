#include "util.h"

#include <algorithm>
#include <cwctype>
#include <format>

namespace rethink {

std::wstring widen(std::string_view utf8) {
    if (utf8.empty()) return {};
    const int n = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, utf8.data(), static_cast<int>(utf8.size()), nullptr, 0);
    if (n <= 0) throw Error("Ungültige UTF-8-Daten");
    std::wstring out(static_cast<size_t>(n), L'\0');
    MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, utf8.data(), static_cast<int>(utf8.size()), out.data(), n);
    return out;
}

std::string narrow(std::wstring_view wide) {
    if (wide.empty()) return {};
    const int n = WideCharToMultiByte(CP_UTF8, 0, wide.data(), static_cast<int>(wide.size()), nullptr, 0, nullptr, nullptr);
    std::string out(static_cast<size_t>(n), '\0');
    WideCharToMultiByte(CP_UTF8, 0, wide.data(), static_cast<int>(wide.size()), out.data(), n, nullptr, nullptr);
    return out;
}

void throwHr(HRESULT hr, std::string_view what) {
    wchar_t* buffer = nullptr;
    const DWORD len = FormatMessageW(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, nullptr,
        static_cast<DWORD>(hr), 0, reinterpret_cast<wchar_t*>(&buffer), 0, nullptr);
    std::string system;
    if (len && buffer) {
        system = narrow(std::wstring_view(buffer, len));
        LocalFree(buffer);
        while (!system.empty() && (system.back() == '\n' || system.back() == '\r' || system.back() == ' ')) system.pop_back();
    }
    throw Error(std::format("{} (HRESULT 0x{:08X}{}{})", what, static_cast<unsigned>(hr), system.empty() ? "" : ": ", system));
}

std::wstring toLower(std::wstring_view s) {
    std::wstring out(s);
    std::transform(out.begin(), out.end(), out.begin(), [](wchar_t c) { return static_cast<wchar_t>(std::towlower(c)); });
    return out;
}

std::wstring xmlEscape(std::wstring_view s) {
    std::wstring out;
    out.reserve(s.size());
    for (wchar_t c : s) {
        switch (c) {
            case L'&': out += L"&amp;"; break;
            case L'<': out += L"&lt;"; break;
            case L'>': out += L"&gt;"; break;
            case L'"': out += L"&quot;"; break;
            case L'\'': out += L"&apos;"; break;
            default: out += c;
        }
    }
    return out;
}

std::wstring xmlEscapeText(std::wstring_view s) {
    std::wstring out;
    out.reserve(s.size());
    for (wchar_t c : s) {
        switch (c) {
            case L'&': out += L"&amp;"; break;
            case L'<': out += L"&lt;"; break;
            case L'>': out += L"&gt;"; break;
            default: out += c;
        }
    }
    return out;
}

std::wstring formatNumber(double value) {
    if (value == 0.0) return L"0";  // vermeidet "-0"
    return std::format(L"{}", value);
}

std::vector<std::uint8_t> readStream(IStream* stream) {
    LARGE_INTEGER zero{};
    check(stream->Seek(zero, STREAM_SEEK_SET, nullptr), "Stream positionieren");
    std::vector<std::uint8_t> out;
    std::vector<std::uint8_t> buffer(1 << 16);
    for (;;) {
        ULONG got = 0;
        const HRESULT hr = stream->Read(buffer.data(), static_cast<ULONG>(buffer.size()), &got);
        if (FAILED(hr)) throwHr(hr, "Stream lesen");
        if (got == 0) break;
        out.insert(out.end(), buffer.begin(), buffer.begin() + got);
    }
    return out;
}

void writeStream(IStream* stream, const std::vector<std::uint8_t>& bytes) {
    ULARGE_INTEGER zeroSize{};
    check(stream->SetSize(zeroSize), "Stream leeren");
    LARGE_INTEGER zero{};
    check(stream->Seek(zero, STREAM_SEEK_SET, nullptr), "Stream positionieren");
    size_t offset = 0;
    while (offset < bytes.size()) {
        const ULONG chunk = static_cast<ULONG>(std::min<size_t>(bytes.size() - offset, 1u << 20));
        ULONG written = 0;
        check(stream->Write(bytes.data() + offset, chunk, &written), "Stream schreiben");
        if (written == 0) throw Error("Stream schreiben: keine Bytes geschrieben");
        offset += written;
    }
}

}  // namespace rethink
