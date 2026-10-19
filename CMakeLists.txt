cmake_minimum_required(VERSION 3.20)
project(rethink VERSION 0.2.0 LANGUAGES CXX RC)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

set(RETHINK_CXX_OPTIONS /W4 /permissive- /utf-8 /EHsc)
set(RETHINK_DEFINES UNICODE _UNICODE NOMINMAX)

# Kern aus Phase 1: von Kommandozeile und Add-in gemeinsam genutzt
add_library(rethink_core STATIC
    src/util.cpp
    src/xmldoc.cpp
    src/opc.cpp
    src/ppttc.cpp
    src/xlsx.cpp
    src/fill_text.cpp
    src/fill_chart.cpp
    src/build.cpp
    src/render_element.cpp
    src/datasheet.cpp
    src/cellinput.cpp
    # Layout-Kern für gezeichnete Diagramme (ohne COM/XML), siehe docs/charting-ledger.md
    src/engine/waterfall.cpp
    src/engine/layout.cpp
    src/engine/dates.cpp
    src/engine/gantt.cpp
    src/engine/gantt_edit.cpp
)
target_include_directories(rethink_core PUBLIC src)
target_compile_definitions(rethink_core PUBLIC ${RETHINK_DEFINES})
target_compile_options(rethink_core PRIVATE ${RETHINK_CXX_OPTIONS})
target_link_libraries(rethink_core PUBLIC ole32 oleaut32 urlmon windowsapp)

# Phase 1: Kommandozeile
add_executable(rethink src/main.cpp)
target_compile_definitions(rethink PRIVATE RETHINK_VERSION="${PROJECT_VERSION}")
target_compile_options(rethink PRIVATE ${RETHINK_CXX_OPTIONS})
target_link_libraries(rethink PRIVATE rethink_core)

# Phase 2: COM-Add-in für PowerPoint
configure_file(addin/version.h.in ${CMAKE_CURRENT_BINARY_DIR}/version.h @ONLY)
add_library(rethink_addin SHARED
    addin/addin.cpp
    addin/commands.cpp
    addin/dispatch.cpp
    addin/elements.cpp
    addin/events.cpp
    addin/overlay.cpp
    addin/datasheet_window.cpp
    addin/calendar_window.cpp
    addin/gallery.cpp
    addin/gantt_menu.cpp
    addin/radial_menu.cpp
    addin/ribbon_icons.cpp
    addin/category_dialog.cpp
    addin/addin.rc
    addin/addin.def
)
set_target_properties(rethink_addin PROPERTIES OUTPUT_NAME rethink-addin)
# Das Ribbon-XML wird per RCDATA eingebettet; Änderungen daran müssen die Ressource neu bauen.
set_source_files_properties(addin/addin.rc PROPERTIES OBJECT_DEPENDS ${CMAKE_CURRENT_SOURCE_DIR}/addin/ribbon.xml)
target_include_directories(rethink_addin PRIVATE addin ${CMAKE_CURRENT_BINARY_DIR})
target_compile_options(rethink_addin PRIVATE $<$<COMPILE_LANGUAGE:CXX>:${RETHINK_CXX_OPTIONS} /WX>)
target_link_libraries(rethink_addin PRIVATE rethink_core user32 gdi32 gdiplus shell32 comdlg32 comctl32 uuid)

# Selbsttest des Add-ins (ohne Registrierung und ohne PowerPoint)
add_executable(addin_selftest tests/addin_selftest.cpp)
target_include_directories(addin_selftest PRIVATE addin)
target_compile_definitions(addin_selftest PRIVATE ${RETHINK_DEFINES})
target_compile_options(addin_selftest PRIVATE ${RETHINK_CXX_OPTIONS} /WX)
target_link_libraries(addin_selftest PRIVATE ole32 oleaut32)

# Unit-Tests des Layout-Kerns
add_executable(engine_test tests/engine_test.cpp)
target_compile_options(engine_test PRIVATE ${RETHINK_CXX_OPTIONS})
target_link_libraries(engine_test PRIVATE rethink_core)
