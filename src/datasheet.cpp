#include "datasheet.h"

#include "util.h"
#include "xlsx.h"

#include <cmath>

namespace rethink {

std::vector<std::uint8_t> datasheetFromTable(const std::vector<Row>& table) {
    std::vector<SheetRow> rows;
    for (auto& row : table) {
        SheetRow sheetRow;
        for (auto& cell : row) {
            SheetCell s;
            s.fill = cell.fill;
            switch (cell.kind) {
                case CellKind::String: s.kind = SheetCell::Kind::Text, s.text = cell.text; break;
                case CellKind::Number: s.kind = SheetCell::Kind::Number, s.value = cell.number; break;
                case CellKind::Percentage: s.kind = SheetCell::Kind::Percent, s.value = cell.number / 100.0; break;
                case CellKind::Date: s.kind = SheetCell::Kind::Date, s.value = excelSerialFromIso(cell.text); break;
                case CellKind::Empty: break;
            }
            sheetRow.push_back(s);
        }
        rows.push_back(std::move(sheetRow));
    }
    return makeWorkbook(rows, L"Datenblatt");
}

std::vector<Row> tableFromDatasheet(const std::vector<std::uint8_t>& bytes) {
    std::vector<Row> table;
    for (auto& sheetRow : readWorkbook(bytes)) {
        Row row;
        for (auto& s : sheetRow) {
            Cell cell;
            cell.fill = s.fill;
            switch (s.kind) {
                case SheetCell::Kind::Text:
                    if (!s.text.empty()) cell.kind = CellKind::String, cell.text = s.text;
                    break;
                case SheetCell::Kind::Number: cell.kind = CellKind::Number, cell.number = s.value; break;
                // Excel speichert 48 % als 0,48; rethink rechnet in Prozentpunkten. Rundung gegen 0,1+0,2-Effekte.
                case SheetCell::Kind::Percent: cell.kind = CellKind::Percentage, cell.number = std::round(s.value * 100.0 * 1e9) / 1e9; break;
                case SheetCell::Kind::Date: cell.kind = CellKind::Date, cell.text = isoFromExcelSerial(s.value); break;
                case SheetCell::Kind::Empty: break;
            }
            row.push_back(cell);
        }
        table.push_back(std::move(row));
    }
    if (table.empty()) table.push_back(Row{});
    return table;
}

std::vector<std::uint8_t> readFileShared(const std::filesystem::path& file) {
    HANDLE h = CreateFileW(file.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING,
                           FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE) throwHr(HRESULT_FROM_WIN32(GetLastError()), "Datenblatt lesen: " + narrow(file.wstring()));
    std::vector<std::uint8_t> bytes;
    LARGE_INTEGER size{};
    if (GetFileSizeEx(h, &size) && size.QuadPart > 0) {
        bytes.resize(static_cast<size_t>(size.QuadPart));
        DWORD read = 0;
        const BOOL ok = ReadFile(h, bytes.data(), static_cast<DWORD>(bytes.size()), &read, nullptr);
        bytes.resize(ok ? read : 0);
    }
    CloseHandle(h);
    return bytes;
}

}  // namespace rethink
