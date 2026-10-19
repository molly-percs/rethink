#include "ribbon_icons.h"

#include <windows.h>
#include <olectl.h>

#include <algorithm>
#include <cmath>
#include <cstring>
#include <initializer_list>
#include <unordered_map>

// GDI+-Köpfe brauchen min/max aus dem globalen Namensraum
using std::max;
using std::min;
#include <gdiplus.h>

namespace rethink::addin {
namespace {

using namespace Gdiplus;

// Palette des Tabs (ARGB)
const Color kBlue(255, 0x25, 0x78, 0xEB), kDark(255, 0x1B, 0x5C, 0xB8), kLight(255, 0xCF, 0xE0, 0xFB), kInk(255, 0x40, 0x40, 0x40),
    kGrey(255, 0x8A, 0x8A, 0x8A), kWhite(255, 255, 255, 255), kGreen(255, 0x2E, 0x9E, 0x5B), kOrange(255, 0xE8, 0x83, 0x3A);

// Zeichner: alle Symbole sind auf einem 16x16-Raster entworfen; die Graphics-Skalierung bringt sie auf 16/32 px (mal DPI),
// eine Linienbreite von 1 Einheit entspricht dann size/16.
struct Pen2D {
    Graphics& g;
    void fill(Color c, REAL x, REAL y, REAL w, REAL h, REAL r = 0) {
        SolidBrush b(c);
        if (r <= 0) return void(g.FillRectangle(&b, x, y, w, h));
        GraphicsPath p;
        round(p, x, y, w, h, r);
        g.FillPath(&b, &p);
    }
    void frame(Color c, REAL x, REAL y, REAL w, REAL h, REAL r = 0, REAL lw = 1, Color inner = Color(0, 0, 0, 0)) {
        GraphicsPath p;
        round(p, x, y, w, h, r);
        if (inner.GetAlpha()) {
            SolidBrush b(inner);
            g.FillPath(&b, &p);
        }
        Pen pen(c, lw);
        pen.SetLineJoin(LineJoinRound);
        g.DrawPath(&pen, &p);
    }
    void line(Color c, REAL x1, REAL y1, REAL x2, REAL y2, REAL lw = 1, bool round = false) {
        Pen pen(c, lw);
        if (round) pen.SetStartCap(LineCapRound), pen.SetEndCap(LineCapRound);
        g.DrawLine(&pen, x1, y1, x2, y2);
    }
    void lines(Color c, std::initializer_list<PointF> pts, REAL lw) {  // offener Linienzug mit runden Enden
        Pen pen(c, lw);
        pen.SetLineJoin(LineJoinRound), pen.SetStartCap(LineCapRound), pen.SetEndCap(LineCapRound);
        g.DrawLines(&pen, pts.begin(), static_cast<INT>(pts.size()));
    }
    void poly(Color c, std::initializer_list<PointF> pts, Color edge = Color(0, 0, 0, 0), REAL lw = 1) {
        SolidBrush b(c);
        g.FillPolygon(&b, pts.begin(), static_cast<INT>(pts.size()));
        if (edge.GetAlpha()) {
            Pen pen(edge, lw);
            pen.SetLineJoin(LineJoinRound);
            g.DrawPolygon(&pen, pts.begin(), static_cast<INT>(pts.size()));
        }
    }
    void pie(Color c, REAL x, REAL y, REAL w, REAL h, REAL start, REAL sweep) {
        SolidBrush b(c);
        g.FillPie(&b, x, y, w, h, start, sweep);
    }
    void disc(Color c, REAL cx, REAL cy, REAL r) {
        SolidBrush b(c);
        g.FillEllipse(&b, cx - r, cy - r, 2 * r, 2 * r);
    }
    void ring(Color c, REAL cx, REAL cy, REAL r, REAL lw = 1) {
        Pen pen(c, lw);
        g.DrawEllipse(&pen, cx - r, cy - r, 2 * r, 2 * r);
    }
    static void round(GraphicsPath& p, REAL x, REAL y, REAL w, REAL h, REAL r) {
        if (r <= 0) {
            p.AddRectangle(RectF(x, y, w, h));
            return;
        }
        const REAL d = 2 * r;
        p.AddArc(x, y, d, d, 180, 90);
        p.AddArc(x + w - d, y, d, d, 270, 90);
        p.AddArc(x + w - d, y + h - d, d, d, 0, 90);
        p.AddArc(x, y + h - d, d, d, 90, 90);
        p.CloseFigure();
    }
    // Seite mit umgeknickter Ecke oben rechts
    void page(REAL x, REAL y, REAL w, REAL h, Color edge, Color fillc) {
        const REAL f = min(w, h) * 0.35f;
        poly(fillc, {{x, y}, {x + w - f, y}, {x + w, y + f}, {x + w, y + h}, {x, y + h}}, edge);
        poly(kLight, {{x + w - f, y}, {x + w - f, y + f}, {x + w, y + f}}, edge);
    }
    // Gefüllte Pfeilspitze in Richtung (dx,dy) mit Spitze bei (x,y)
    void head(Color c, REAL x, REAL y, REAL dx, REAL dy, REAL len, REAL half) {
        const REAL bx = x - dx * len, by = y - dy * len;
        poly(c, {{x, y}, {bx - dy * half, by + dx * half}, {bx + dy * half, by - dx * half}});
    }
};

using Drawing = void (*)(Pen2D&);

// Tabelle id -> Zeichnung. Große Symbole (Gruppenköpfe) zuerst, danach die kleinen Zeilen-Symbole.
const std::unordered_map<std::wstring, Drawing>& drawings() {
    static const std::unordered_map<std::wstring, Drawing> t = {
        {L"agenda", [](Pen2D& d) {
             for (int i = 0; i < 3; ++i) {
                 const REAL y = 1.5f + i * 5;
                 d.fill(kBlue, 1, y, 4, 3.5f, 0.8f);
                 d.line(kGrey, 7, y + 1.75f, 15, y + 1.75f, 1.6f);
             }
         }},
        {L"library", [](Pen2D& d) {
             d.frame(kGrey, 5.5f, 1.5f, 9, 6.5f, 1, 1, kWhite);
             d.frame(kGrey, 3.25f, 4.5f, 9, 6.5f, 1, 1, kWhite);
             d.frame(kDark, 1.5f, 7.5f, 9, 7, 1, 1, kWhite);
             d.fill(kBlue, 1.5f, 7.5f, 9, 2.5f, 1);
             d.fill(kBlue, 1.5f, 8.5f, 9, 1.5f);
         }},
        {L"charts", [](Pen2D& d) {
             d.fill(kBlue, 1.5f, 8, 3.5f, 6, 0.6f);
             d.fill(kDark, 6.25f, 2.5f, 3.5f, 11.5f, 0.6f);
             d.fill(kOrange, 11, 5.5f, 3.5f, 8.5f, 0.6f);
             d.line(kGrey, 0.75f, 14.75f, 15.25f, 14.75f, 1.2f);
         }},
        {L"layout", [](Pen2D& d) {
             for (int i = 0; i < 4; ++i)
                 d.frame(kGrey, 1.5f + (i & 1) * 7.5f, 1.5f + (i >> 1) * 7.5f, 5.5f, 5.5f, 1.2f, 1, kLight);
             d.line(kDark, 8, 0.5f, 8, 15.5f, 1.4f);
         }},
        {L"color", [](Pen2D& d) {
             d.poly(kInk, {{1.5f, 6.5f}, {6.5f, 1.5f}, {11, 6}, {6, 11}}, kInk, 0.6f);  // Eimer
             d.lines(kGrey, {{4.5f, 4}, {3, 2.5f}, {5, 1}}, 1);                       // Henkel
             d.line(kWhite, 3.5f, 7, 7, 3.5f, 0.9f);
             d.poly(kBlue, {{13, 7}, {14.6f, 9.8f}, {11.4f, 9.8f}});                  // Tropfen
             d.disc(kBlue, 13, 10.2f, 1.7f);
             d.fill(kBlue, 1, 12.5f, 14, 3, 0.8f);
         }},
        {L"quality", [](Pen2D& d) {
             d.page(2.5f, 1, 11, 14, kGrey, kWhite);
             d.lines(kGreen, {{5, 9.5f}, {7.3f, 12}, {12, 6.5f}}, 2);
         }},
        {L"collab", [](Pen2D& d) {
             d.disc(kGrey, 10.8f, 4.3f, 2.4f);
             d.pie(kGrey, 6.8f, 8, 8.4f, 9, 180, 180);
             d.disc(kWhite, 6, 6.2f, 3.2f);  // heller Rand trennt die vordere Figur
             d.pie(kWhite, 0.2f, 9.4f, 11.6f, 11, 180, 180);
             d.disc(kBlue, 6, 6.2f, 2.5f);
             d.pie(kBlue, 1, 10.2f, 10, 9.4f, 180, 180);
         }},
        {L"more", [](Pen2D& d) {
             d.ring(kBlue, 8, 8, 6.8f, 1.4f);
             for (int i = -1; i <= 1; ++i) d.disc(kBlue, 8 + i * 3.4f, 8, 1.1f);
         }},
        {L"textbox", [](Pen2D& d) {
             d.lines(kInk, {{1.5f, 14}, {5.5f, 2}, {9.5f, 14}}, 2);
             d.line(kInk, 3, 10.3f, 8, 10.3f, 1.8f);
             d.line(kBlue, 11.5f, 6, 15, 6, 1.6f);
             d.line(kBlue, 11.5f, 10, 15, 10, 1.6f);
         }},
        {L"process", [](Pen2D& d) { d.poly(kBlue, {{1, 3.5f}, {10, 3.5f}, {15, 8}, {10, 12.5f}, {1, 12.5f}}); }},
        {L"roundrect", [](Pen2D& d) { d.fill(kBlue, 1, 3, 14, 10, 3); }},
        {L"connector", [](Pen2D& d) {
             d.line(kGrey, 3.5f, 12.5f, 12.5f, 3.5f, 1.4f);
             d.disc(kBlue, 3.5f, 12.5f, 2.2f);
             d.disc(kBlue, 12.5f, 3.5f, 2.2f);
         }},
        {L"harvey", [](Pen2D& d) {
             d.pie(kBlue, 1.5f, 1.5f, 13, 13, -90, 90);
             d.ring(kBlue, 8, 8, 6.5f, 1.3f);
         }},
        {L"checkbox", [](Pen2D& d) {
             d.frame(kGrey, 1.5f, 1.5f, 13, 13, 2, 1.2f);
             d.lines(kBlue, {{4.5f, 8.2f}, {7.0f, 10.8f}, {11.5f, 5.0f}}, 1.9f);
         }},
        {L"list", [](Pen2D& d) {
             for (int i = 0; i < 3; ++i) {
                 const REAL y = 3.0f + i * 5.0f;
                 d.disc(kBlue, 2.7f, y, 1.5f);
                 d.line(kGrey, 6, y, 15, y, 1.6f);
             }
         }},
        {L"table", [](Pen2D& d) {
             d.frame(kGrey, 1.5f, 2.5f, 13, 11, 0, 1, kWhite);
             d.fill(kBlue, 1.5f, 2.5f, 13, 3.7f);
             for (int i = 1; i <= 2; ++i) d.line(kGrey, 1.5f, 2.5f + 3.67f * i, 14.5f, 2.5f + 3.67f * i);
             d.line(kGrey, 5.8f, 6.2f, 5.8f, 13.5f);
             d.line(kGrey, 10.2f, 6.2f, 10.2f, 13.5f);
         }},
        {L"shapes", [](Pen2D& d) {
             d.fill(kGrey, 1.5f, 7, 8, 8);
             d.disc(kWhite, 10, 6.5f, 5.6f);
             d.disc(kBlue, 10, 6.5f, 4.8f);
         }},
        {L"images", [](Pen2D& d) {
             d.frame(kGrey, 1.5f, 2.5f, 13, 11, 1, 1, kWhite);
             d.poly(kDark, {{2.5f, 12.5f}, {6, 7}, {8.7f, 10.5f}, {10.5f, 8.5f}, {13.5f, 12.5f}});
             d.disc(kBlue, 11.2f, 5.4f, 1.5f);
         }},
        {L"datasheet", [](Pen2D& d) {
             d.frame(kGrey, 1.5f, 1.5f, 13, 13, 0, 1, kWhite);
             d.fill(kLight, 1.5f, 4.5f, 4, 10);
             d.fill(kBlue, 1.5f, 1.5f, 13, 3);
             for (int i = 1; i <= 3; ++i) d.line(kGrey, 1.5f, 1.5f + 3.25f * i, 14.5f, 1.5f + 3.25f * i);
             d.line(kGrey, 5.5f, 4.5f, 5.5f, 14.5f);
             d.line(kGrey, 10, 4.5f, 10, 14.5f);
         }},
        {L"workbook", [](Pen2D& d) {
             d.frame(kGrey, 1.5f, 1.5f, 11, 11, 0, 1, kWhite);
             d.fill(kGreen, 1.5f, 1.5f, 11, 3);
             d.line(kGrey, 1.5f, 7.5f, 12.5f, 7.5f);
             d.line(kGrey, 5.2f, 4.5f, 5.2f, 12.5f);
             d.line(kGrey, 8.8f, 4.5f, 8.8f, 12.5f);
             d.poly(kBlue, {{7.5f, 11}, {10.5f, 11}, {10.5f, 9.5f}, {15, 12.5f}, {10.5f, 15.5f}, {10.5f, 14}, {7.5f, 14}}, kWhite, 1);
         }},
        {L"ppttc", [](Pen2D& d) {
             d.page(5, 1, 10, 14, kGrey, kWhite);
             d.poly(kBlue, {{0.8f, 6.5f}, {5, 6.5f}, {5, 4.5f}, {9, 8.5f}, {5, 12.5f}, {5, 10.5f}, {0.8f, 10.5f}}, kWhite, 1);
         }},
        {L"refresh", [](Pen2D& d) {
             Pen pen(kBlue, 1.9f);
             d.g.DrawArc(&pen, 2.5f, 2.5f, 11.0f, 11.0f, 35.0f, 280.0f);
             const REAL a = -45 * 3.14159265f / 180, c = std::cos(a), s = std::sin(a);  // Bogenende bei -45 Grad, im Uhrzeigersinn
             d.head(kBlue, 8 + 5.5f * c - s * 2.6f, 8 + 5.5f * s + c * 2.6f, -s, c, 2.6f, 2.8f);
         }},
        {L"link", [](Pen2D& d) {
             GraphicsPath p(FillModeAlternate);
             const PointF tag[] = {{1, 3}, {10, 3}, {15, 8}, {10, 13}, {1, 13}};
             p.AddPolygon(tag, 5);
             p.AddEllipse(8.8f, 6.8f, 2.4f, 2.4f);
             SolidBrush b(kBlue);
             d.g.FillPath(&b, &p);
             d.line(kWhite, 3, 6.5f, 7, 6.5f, 1.2f);
             d.line(kWhite, 3, 9.5f, 6, 9.5f, 1.2f);
         }},
        {L"excellink", [](Pen2D& d) {
             d.frame(kGrey, 1.5f, 1.5f, 9.5f, 9.5f, 0, 1, kWhite);
             d.fill(kGreen, 1.5f, 1.5f, 9.5f, 3);
             d.line(kGrey, 1.5f, 7.8f, 11, 7.8f);
             d.line(kGrey, 6.2f, 4.5f, 6.2f, 11);
             for (int i = 0; i < 2; ++i) {  // zwei verhakte Kettenglieder, um 45 Grad gedreht
                 GraphicsState st = d.g.Save();
                 d.g.TranslateTransform(11 + (i ? 1.7f : -1.7f), 11 + (i ? -1.7f : 1.7f));
                 d.g.RotateTransform(-45);
                 d.frame(kBlue, -3.2f, -1.8f, 6.4f, 3.6f, 1.8f, 1.5f, i ? Color(0, 0, 0, 0) : kWhite);
                 d.g.Restore(st);
             }
         }},
        {L"autofit", [](Pen2D& d) {
             d.frame(kGrey, 1.5f, 1.5f, 13, 8, 1, 1.1f);
             d.line(kInk, 3.5f, 4.2f, 12.5f, 4.2f, 1.2f);
             d.line(kInk, 3.5f, 6.8f, 9.5f, 6.8f, 1.2f);
             d.line(kBlue, 3.5f, 13, 12.5f, 13, 1.4f);
             d.head(kBlue, 1, 13, -1, 0, 3, 2.2f);
             d.head(kBlue, 15, 13, 1, 0, 3, 2.2f);
         }},
        {L"margins", [](Pen2D& d) {
             d.frame(kGrey, 1.5f, 1.5f, 13, 13, 1, 1.2f);
             Pen pen(kBlue, 1.1f);
             pen.SetDashStyle(DashStyleDash);
             d.g.DrawRectangle(&pen, 4.5f, 4.5f, 7.0f, 7.0f);
         }},
        {L"paragraph", [](Pen2D& d) {
             d.pie(kInk, 2.5f, 1.5f, 11, 9, 90, 180);
             d.fill(kInk, 8, 1.5f, 1.8f, 13);
             d.fill(kInk, 11.2f, 1.5f, 1.8f, 13);
             d.fill(kInk, 8, 1.5f, 5, 1.6f);
         }},
    };
    return t;
}

// GDI+ einmalig starten; im geladenen DLL-Leben nie beenden (Aufräumen aus DllMain wäre unzulässig)
bool startGdiplus() {
    static const bool ok = [] {
        static ULONG_PTR token = 0;
        const GdiplusStartupInput in;
        return GdiplusStartup(&token, &in, nullptr) == Ok;
    }();
    return ok;
}

}  // namespace

IPictureDisp* ribbonIcon(const std::wstring& id, bool large) {
    const auto it = drawings().find(id);
    if (it == drawings().end() || !startGdiplus()) return nullptr;

    const int px = static_cast<int>(std::lround((large ? 32 : 16) * (GetDpiForSystem() / 96.0)));
    Bitmap bmp(px, px, PixelFormat32bppARGB);
    {
        Graphics g(&bmp);
        g.SetSmoothingMode(SmoothingModeAntiAlias);
        g.SetPixelOffsetMode(PixelOffsetModeHalf);
        g.ScaleTransform(px / 16.0f, px / 16.0f);
        Pen2D pen{g};
        it->second(pen);
    }

    // Office zeigt 32-bpp-DIBs mit vorbelegtem Alpha transparent an
    BitmapData data{};
    const Rect all(0, 0, px, px);
    if (bmp.LockBits(&all, ImageLockModeRead, PixelFormat32bppPARGB, &data) != Ok) return nullptr;
    BITMAPINFO bi{};
    bi.bmiHeader = {sizeof(BITMAPINFOHEADER), px, -px, 1, 32, BI_RGB, 0, 0, 0, 0, 0};
    void* bits = nullptr;
    HBITMAP hbm = CreateDIBSection(nullptr, &bi, DIB_RGB_COLORS, &bits, nullptr, 0);
    if (hbm && bits)
        for (int y = 0; y < px; ++y)
            std::memcpy(static_cast<char*>(bits) + static_cast<size_t>(y) * px * 4, static_cast<char*>(data.Scan0) + static_cast<ptrdiff_t>(y) * data.Stride,
                        static_cast<size_t>(px) * 4);
    bmp.UnlockBits(&data);
    if (!hbm) return nullptr;

    PICTDESC desc{sizeof(PICTDESC), PICTYPE_BITMAP};
    desc.bmp.hbitmap = hbm;
    IPictureDisp* pic = nullptr;
    if (FAILED(OleCreatePictureIndirect(&desc, IID_IPictureDisp, TRUE, reinterpret_cast<void**>(&pic)))) {
        DeleteObject(hbm);
        return nullptr;
    }
    return pic;
}

}  // namespace rethink::addin
