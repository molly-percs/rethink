#include <winrt/base.h>

#include "build.h"
#include "util.h"

#include <cstdio>
#include <filesystem>
#include <set>
#include <string>
#include <vector>

namespace {

void printUsage() {
    std::fputs(
        "rethink " RETHINK_VERSION " - Präsentationen aus .ppttc-Daten erzeugen\n"
        "\n"
        "Aufruf:\n"
        "  rethink build <eingabe.ppttc> [-o <ausgabe.pptx>]\n"
        "  rethink --version\n"
        "\n"
        "Elemente der Vorlage werden über ihren Namen gefunden:\n"
        "  - Formname im Auswahlbereich oder Alternativtext \"rethink:<Name>\"\n"
        "  - Textfelder aus think-cell-Vorlagen (kompatibel)\n",
        stdout);
}

int usageError(const char* message) {
    std::fprintf(stderr, "Fehler: %s\n\n", message);
    printUsage();
    return 1;
}

}  // namespace

int wmain(int argc, wchar_t** argv) {
    SetConsoleOutputCP(CP_UTF8);
    const std::vector<std::wstring> args(argv + 1, argv + argc);

    if (args.empty()) return usageError("Kein Befehl angegeben");
    if (args[0] == L"--help" || args[0] == L"-h" || args[0] == L"help") {
        printUsage();
        return 0;
    }
    if (args[0] == L"--version") {
        std::puts("rethink " RETHINK_VERSION);
        return 0;
    }
    if (args[0] != L"build") return usageError("Unbekannter Befehl");

    std::filesystem::path input, output;
    for (size_t i = 1; i < args.size(); ++i) {
        if (args[i] == L"-o" || args[i] == L"--output") {
            if (++i >= args.size()) return usageError("Nach -o fehlt der Dateiname");
            output = args[i];
        } else if (input.empty()) {
            input = args[i];
        } else {
            return usageError("Zu viele Argumente");
        }
    }
    if (input.empty()) return usageError("Keine .ppttc-Datei angegeben");
    if (output.empty()) output = std::filesystem::path(input).replace_extension(L".pptx");

    try {
        winrt::init_apartment(winrt::apartment_type::single_threaded);
        const auto result = rethink::build(input, output);
        std::set<std::string> shown;  // gleiche Warnung aus mehreren Einträgen nur einmal zeigen
        for (auto& w : result.warnings)
            if (shown.insert(w).second) std::fprintf(stderr, "Warnung: %s\n", w.c_str());
        std::printf("%s: %d Folien, %d Elemente gefüllt\n", rethink::narrow(output.wstring()).c_str(), result.slides, result.filled);
        return 0;
    } catch (const rethink::Error& e) {
        std::fprintf(stderr, "Fehler: %s\n", e.what());
        return 2;
    } catch (const winrt::hresult_error& e) {
        std::fprintf(stderr, "Fehler: %s\n", rethink::narrow(e.message()).c_str());
        return 3;
    } catch (const std::exception& e) {
        std::fprintf(stderr, "Fehler: %s\n", e.what());
        return 3;
    }
}
