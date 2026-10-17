#include <winrt/base.h>

#include "commands.h"

#include "build.h"
#include "dispatch.h"
#include "resource.h"

#include <shobjidl.h>

#include <exception>
#include <format>
#include <set>
#include <thread>

extern "C" IMAGE_DOS_HEADER __ImageBase;

namespace rethink::addin {
namespace fs = std::filesystem;
using disp::Value;

namespace {

constexpr wchar_t kTitle[] = L"rethink";
constexpr wchar_t kPrefix[] = L"rethink:";
constexpr long kXlColumnClustered = 51;
constexpr long kPpSelectionShapes = 2;
constexpr long kPpSelectionText = 3;

HINSTANCE moduleHandle() { return reinterpret_cast<HINSTANCE>(&__ImageBase); }

HWND ownerWindow(IDispatch* app) {
    try {
        return reinterpret_cast<HWND>(static_cast<LONG_PTR>(disp::call(app, L"HWND").toLong()));
    } catch (...) {
        return GetActiveWindow();
    }
}

int message(HWND owner, const std::wstring& text, UINT flags) { return MessageBoxW(owner, text.c_str(), kTitle, flags); }

// Führt einen Ribbon-Befehl aus; Fehler erscheinen als Meldung statt die COM-Grenze zu überqueren.
template <class F>
void guarded(IDispatch* app, F&& body) {
    HWND owner = nullptr;
    try {
        owner = ownerWindow(app);
        body(owner);
        return;
    } catch (const winrt::hresult_error& e) {
        message(owner, e.message().c_str(), MB_ICONERROR);
    } catch (const std::exception& e) {
        std::wstring text;
        try {
            text = widen(e.what());
        } catch (...) {
            text = L"Unerwarteter Fehler";
        }
        message(owner, text, MB_ICONERROR);
    } catch (...) {
        message(owner, L"Unerwarteter Fehler", MB_ICONERROR);
    }
}

std::wstring trim(std::wstring_view s) {
    const auto first = s.find_first_not_of(L" \t\r\n");
    if (first == std::wstring_view::npos) return {};
    const auto last = s.find_last_not_of(L" \t\r\n");
    return std::wstring(s.substr(first, last - first + 1));
}

std::wstring normalized(const fs::path& p) {
    std::error_code ec;
    auto canonical = fs::weakly_canonical(p, ec);
    return toLower((ec ? p : canonical).wstring());
}

bool isOpenInPowerPoint(IDispatch* app, const fs::path& file) {
    auto presentations = disp::object(app, L"Presentations");
    const long count = disp::call(presentations.Get(), L"Count").toLong();
    const auto wanted = normalized(file);
    for (long i = 1; i <= count; ++i) {
        auto pres = disp::object(presentations.Get(), L"Item", i);
        if (normalized(disp::call(pres.Get(), L"FullName").str()) == wanted) return true;
    }
    return false;
}

bool isTaken(IDispatch* app, const fs::path& file) { return fs::exists(file) || isOpenInPowerPoint(app, file); }

fs::path freeOutputName(IDispatch* app, const fs::path& wanted) {
    for (int n = 2;; ++n) {
        auto candidate = wanted.parent_path() / std::format(L"{} ({}){}", wanted.stem().wstring(), n, wanted.extension().wstring());
        if (!isTaken(app, candidate)) return candidate;
    }
}

// rethink::build läuft in einem eigenen Thread mit eigenem Apartment, genau wie im Kommandozeilenwerkzeug.
// So bleibt der PowerPoint-Thread unberührt (kein init_apartment im fremden STA).
BuildResult runBuild(const fs::path& input, const fs::path& output) {
    BuildResult result;
    std::exception_ptr error;
    std::thread worker([&] {
        try {
            winrt::init_apartment(winrt::apartment_type::single_threaded);
            try {
                result = build(input, output);
            } catch (...) {
                error = std::current_exception();
            }
            winrt::uninit_apartment();
        } catch (...) {
            error = std::current_exception();
        }
    });
    worker.join();
    if (error) std::rethrow_exception(error);
    return result;
}

ComPtr<IDispatch> activeWindow(IDispatch* app) {
    auto windows = disp::object(app, L"Windows");
    if (disp::call(windows.Get(), L"Count").toLong() < 1) throw Error("Es ist keine Präsentation geöffnet.");
    return disp::object(app, L"ActiveWindow");
}

ComPtr<IDispatch> selectedShapes(IDispatch* app) {
    auto selection = disp::object(activeWindow(app).Get(), L"Selection");
    const long type = disp::call(selection.Get(), L"Type").toLong();
    if (type != kPpSelectionShapes && type != kPpSelectionText) throw Error("Bitte zuerst eine oder mehrere Formen auswählen.");
    bool child = false;
    try {
        child = disp::call(selection.Get(), L"HasChildShapeRange").toLong() != 0;  // Formen innerhalb einer Gruppe
    } catch (const Error&) {
    }
    return disp::object(selection.Get(), child ? L"ChildShapeRange" : L"ShapeRange");
}

std::wstring validName(const std::wstring& raw) {
    std::wstring name = trim(raw);
    if (name.find_first_of(L"\r\n") != std::wstring::npos) throw Error("Der Name darf keinen Zeilenumbruch enthalten.");
    return name;
}

struct SelectionInfo {
    long count = 0;
    std::wstring name;         // vorgeschlagener Name
    std::wstring foreignText;  // vorhandener Alternativtext ohne rethink-Präfix
};

SelectionInfo describeSelection(IDispatch* app) {
    auto range = selectedShapes(app);
    SelectionInfo info;
    info.count = disp::call(range.Get(), L"Count").toLong();
    if (info.count < 1) throw Error("Bitte zuerst eine oder mehrere Formen auswählen.");
    for (long i = 1; i <= info.count; ++i) {
        auto shape = disp::object(range.Get(), L"Item", i);
        const auto alt = disp::call(shape.Get(), L"AlternativeText").str();
        if (alt.starts_with(kPrefix)) {
            if (info.name.empty()) info.name = alt.substr(std::size(kPrefix) - 1);
        } else if (!trim(alt).empty() && info.foreignText.empty()) {
            info.foreignText = alt;
        }
        if (i == 1 && !alt.starts_with(kPrefix)) info.name = disp::call(shape.Get(), L"Name").str();
    }
    return info;
}

// --- Dialoge ---

fs::path pickPpttc(HWND owner) {
    ComPtr<IFileOpenDialog> dialog;
    check(CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&dialog)), "Dateidialog");
    const COMDLG_FILTERSPEC types[] = {{L"rethink-Daten (*.ppttc)", L"*.ppttc"}, {L"Alle Dateien (*.*)", L"*.*"}};
    dialog->SetFileTypes(static_cast<UINT>(std::size(types)), types);
    dialog->SetTitle(L"rethink – .ppttc-Datei wählen");
    dialog->SetOkButtonLabel(L"Erzeugen");
    DWORD options = 0;
    dialog->GetOptions(&options);
    dialog->SetOptions(options | FOS_FORCEFILESYSTEM | FOS_FILEMUSTEXIST);
    const HRESULT hr = dialog->Show(owner);
    if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED)) return {};
    check(hr, "Dateidialog");
    ComPtr<IShellItem> item;
    check(dialog->GetResult(&item), "Dateidialog");
    CoStr path;
    check(item->GetDisplayName(SIGDN_FILESYSPATH, path.put()), "Dateidialog");
    return path.str();
}

struct NameDialog {
    SelectionInfo info;
    std::wstring result;
};

INT_PTR CALLBACK nameDialogProc(HWND dlg, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_INITDIALOG: {
        SetWindowLongPtrW(dlg, DWLP_USER, lParam);
        auto* state = reinterpret_cast<NameDialog*>(lParam);
        SetWindowTextW(dlg, L"rethink – Elementname setzen");
        const auto prompt = state->info.count == 1
                                ? std::wstring(L"Name der ausgewählten Form. Unter diesem Namen füllt rethink sie aus einer .ppttc-Datei:")
                                : std::format(L"Name für die {} ausgewählten Formen. Unter diesem Namen füllt rethink sie aus einer .ppttc-Datei:", state->info.count);
        SetDlgItemTextW(dlg, IDC_PROMPT, prompt.c_str());
        SetDlgItemTextW(dlg, IDC_NAME, state->info.name.c_str());
        const auto hint = state->info.foreignText.empty()
                              ? std::wstring(L"Gespeichert als Alternativtext „rethink:<Name>“. Leer lassen entfernt den Namen.")
                              : std::format(L"Achtung: Der bisherige Alternativtext „{}“ wird ersetzt.", state->info.foreignText.substr(0, 80));
        SetDlgItemTextW(dlg, IDC_HINT, hint.c_str());
        SetDlgItemTextW(dlg, IDCANCEL, L"Abbrechen");
        HWND edit = GetDlgItem(dlg, IDC_NAME);
        SendMessageW(edit, EM_SETSEL, 0, -1);
        SetFocus(edit);
        return FALSE;  // Fokus selbst gesetzt
    }
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK) {
            auto* state = reinterpret_cast<NameDialog*>(GetWindowLongPtrW(dlg, DWLP_USER));
            HWND edit = GetDlgItem(dlg, IDC_NAME);
            std::wstring text(static_cast<size_t>(GetWindowTextLengthW(edit)) + 1, L'\0');
            text.resize(static_cast<size_t>(GetWindowTextW(edit, text.data(), static_cast<int>(text.size()))));
            state->result = text;
            EndDialog(dlg, IDOK);
            return TRUE;
        }
        if (LOWORD(wParam) == IDCANCEL) {
            EndDialog(dlg, IDCANCEL);
            return TRUE;
        }
        break;
    }
    return FALSE;
}

}  // namespace

// --- Befehle ohne Dialoge ---

std::wstring buildAndOpen(IDispatch* app, const fs::path& input, fs::path output, std::vector<std::string>& warnings) {
    if (output.empty()) {
        output = fs::path(input).replace_extension(L".pptx");
        if (isTaken(app, output)) output = freeOutputName(app, output);
    } else if (isOpenInPowerPoint(app, output)) {
        throw Error("Die Zieldatei ist in PowerPoint geöffnet: " + narrow(output.wstring()));
    }

    const auto result = runBuild(input, output);
    std::set<std::string> shown;  // gleiche Warnung aus mehreren Einträgen nur einmal
    for (auto& w : result.warnings)
        if (shown.insert(w).second) warnings.push_back(w);

    auto presentations = disp::object(app, L"Presentations");
    auto pres = disp::object(presentations.Get(), L"Open", output.wstring());
    return disp::call(pres.Get(), L"FullName").str();
}

std::wstring insertChart(IDispatch* app) {
    auto window = activeWindow(app);
    ComPtr<IDispatch> slide;
    try {
        slide = disp::object(disp::object(window.Get(), L"View").Get(), L"Slide");
    } catch (const Error&) {
        throw Error("Bitte in der Normalansicht eine Folie auswählen.");
    }

    auto setup = disp::object(disp::object(window.Get(), L"Presentation").Get(), L"PageSetup");
    const double slideWidth = disp::call(setup.Get(), L"SlideWidth").toDouble();
    const double slideHeight = disp::call(setup.Get(), L"SlideHeight").toDouble();

    // Freien Namen "Chart<n>" suchen (Formname und Alternativtext).
    auto shapes = disp::object(slide.Get(), L"Shapes");
    std::set<std::wstring> taken;
    const long count = disp::call(shapes.Get(), L"Count").toLong();
    for (long i = 1; i <= count; ++i) {
        auto shape = disp::object(shapes.Get(), L"Item", i);
        taken.insert(toLower(disp::call(shape.Get(), L"Name").str()));
        const auto alt = disp::call(shape.Get(), L"AlternativeText").str();
        if (alt.starts_with(kPrefix)) taken.insert(toLower(alt.substr(std::size(kPrefix) - 1)));
    }
    std::wstring name;
    for (int n = 1; name.empty() || taken.contains(toLower(name)); ++n) name = std::format(L"Chart{}", n);

    const double width = slideWidth * 0.6, height = slideHeight * 0.6;
    auto shape = disp::object(shapes.Get(), L"AddChart2", -1L, kXlColumnClustered, (slideWidth - width) / 2,
                              (slideHeight - height) / 2, width, height);
    try {
        // AddChart2 kann das Datenblatt öffnen; es wird ohne Änderung geschlossen.
        auto data = disp::object(disp::object(shape.Get(), L"Chart").Get(), L"ChartData");
        disp::call(disp::object(data.Get(), L"Workbook").Get(), L"Close");
    } catch (const Error&) {
    }
    disp::put(shape.Get(), L"Name", name);
    disp::put(shape.Get(), L"AlternativeText", std::wstring(kPrefix) + name);
    try {
        disp::call(shape.Get(), L"Select");
    } catch (const Error&) {
    }
    return name;
}

int setElementName(IDispatch* app, const std::wstring& rawName) {
    const std::wstring name = validName(rawName);
    auto range = selectedShapes(app);
    const long count = disp::call(range.Get(), L"Count").toLong();
    int changed = 0;
    for (long i = 1; i <= count; ++i) {
        auto shape = disp::object(range.Get(), L"Item", i);
        if (name.empty()) {
            if (!disp::call(shape.Get(), L"AlternativeText").str().starts_with(kPrefix)) continue;
            disp::put(shape.Get(), L"AlternativeText", L"");
        } else {
            disp::put(shape.Get(), L"AlternativeText", std::wstring(kPrefix) + name);
        }
        ++changed;
    }
    return changed;
}

std::wstring elementName(IDispatch* app) { return describeSelection(app).name; }

// --- Ribbon-Befehle ---

void commandBuild(IDispatch* app) {
    guarded(app, [&](HWND owner) {
        const auto input = pickPpttc(owner);
        if (input.empty()) return;

        auto output = fs::path(input).replace_extension(L".pptx");
        if (isOpenInPowerPoint(app, output)) {
            output = freeOutputName(app, output);
        } else if (fs::exists(output)) {
            const auto answer = message(owner,
                                        std::format(L"„{}“ existiert bereits.\n\nJa: überschreiben\nNein: als „{}“ speichern",
                                                    output.filename().wstring(), freeOutputName(app, output).filename().wstring()),
                                        MB_YESNOCANCEL | MB_ICONQUESTION);
            if (answer == IDCANCEL) return;
            if (answer == IDNO) output = freeOutputName(app, output);
        }

        HCURSOR previous = SetCursor(LoadCursorW(nullptr, IDC_WAIT));
        std::vector<std::string> warnings;
        try {
            buildAndOpen(app, input, output, warnings);
        } catch (...) {
            SetCursor(previous);
            throw;
        }
        SetCursor(previous);

        if (!warnings.empty()) {
            std::wstring text = std::format(L"„{}“ wurde erzeugt. Hinweise:\n", output.filename().wstring());
            for (auto& w : warnings) text += L"\n• " + widen(w);
            message(ownerWindow(app), text, MB_ICONWARNING);
        }
    });
}

void commandInsertChart(IDispatch* app) {
    guarded(app, [&](HWND) { insertChart(app); });
}

void commandSetName(IDispatch* app) {
    guarded(app, [&](HWND owner) {
        NameDialog state{describeSelection(app), {}};
        const INT_PTR answer =
            DialogBoxParamW(moduleHandle(), MAKEINTRESOURCEW(IDD_NAME), owner, nameDialogProc, reinterpret_cast<LPARAM>(&state));
        if (answer != IDOK) return;
        setElementName(app, state.result);
    });
}

}  // namespace rethink::addin
