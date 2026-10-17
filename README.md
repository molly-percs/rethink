# rethink 0.2

`rethink` erzeugt PowerPoint-Präsentationen aus Daten im `.ppttc`-JSON-Format. Das ist das dokumentierte Austauschformat der think-cell-Automatisierung.

Es handelt sich um einen **eigenständigen Clean-Room-Nachbau**. Grundlage waren nur das öffentlich dokumentierte Dateiformat und das beschriebene Verhalten. Es wurde kein Code dekompiliert oder übernommen, und rethink verwendet keine Kennungen, Namen oder Markenzeichen von think-cell. Die Spezifikation der Diagrammfunktionen und ihr Stand stehen in [docs/charting-ledger.md](docs/charting-ledger.md).

> Status: Arbeitsstand in Entwicklung (Windows, PowerPoint/Microsoft 365). Fertige Binärdateien liegen nicht im Repository; siehe „Bauen“.

## Schnell ausprobieren

```powershell
.\ausprobieren.ps1                         # baut samples\sample.ppttc und öffnet das Ergebnis in PowerPoint
.\ausprobieren.ps1 -Ppttc meine-daten.ppttc
```

Werte ändern: `samples\sample.ppttc` in einem Editor öffnen, Zahlen oder Texte anpassen, speichern und das Skript erneut starten.

## Aufruf

```
rethink build <eingabe.ppttc> [-o <ausgabe.pptx>]
rethink --version
```

- Ohne `-o` entsteht `<eingabe>.pptx` neben der Eingabedatei.
- Fehlende Ausgabeordner werden angelegt.
- Exit-Codes: `0` bedeutet Erfolg (Warnungen sind möglich), `1` falscher Aufruf, `2` ungültige Eingabe oder Vorlage, `3` unerwarteter Fehler.

## Was Phase 1 kann

| Funktion | Verhalten |
|---|---|
| `.ppttc` lesen | Validiert das Format vollständig: `template`, `data[]`, `name`, `table`, Zelltypen `string`/`number`/`percentage`/`date`/`null`, `fill` als `#rrggbb` oder `rgb(r,g,b)`. Fehlermeldungen nennen den JSON-Pfad, z. B. `[0].data[3].table[1][2]`. |
| Vorlagen | Lokaler Pfad (relativ zur `.ppttc`-Datei) oder `http(s)`-URL. |
| Folien vervielfältigen | Jeder Eintrag erzeugt eine Kopie **aller** Vorlagenfolien. Diagramme, Einbettungen und Tags werden mitkopiert, Layouts, Master und Bilder geteilt, Notizen weggelassen. |
| Textfelder | Formen werden über den **Formnamen** (Auswahlbereich, Alt+F10) oder den **Alternativtext `rethink:<Name>`** gefunden. Außerdem werden die **Textfelder in think-cell-Vorlagen** erkannt (kompatibel). Die Formatierung des ersten Textlaufs bleibt erhalten, `\n` erzeugt neue Absätze. |
| Diagramme | Native PowerPoint-Diagramme (Säule, Linie, Fläche, Kreis; Balken und Ring ungetestet), gleiche Namensregel wie bei Textfeldern. Zeile 0 enthält die Kategorien, jede weitere Zeile eine Reihe. Eine **leere Zeile `[]` überspringt eine Farbe** des Füllschemas. `percentage` wird mit dem Format `0%` geschrieben, `fill` färbt einzelne Datenpunkte. |
| Datumskategorien | Lauter 1. Januar → `2024`; lauter Monatserste → `2024-03`; sonst das Datum wie angegeben. |
| Datenblatt | Jedes Diagramm bekommt eine neue eingebettete XLSX-Datei mit denselben Werten. So zeigt „Daten bearbeiten“ in PowerPoint die richtigen Zahlen. |
| Ergebnis | Eine normale PPTX-Datei, die **ohne Add-in** geöffnet und bearbeitet werden kann. |

## Gezeichnete Diagramme (think-cell-Stil)

Bestimmte Diagramme kann PowerPoint nativ nicht darstellen: Wasserfälle mit `e`, eine Datentabelle unter dem Diagramm, später Mekko, CAGR-Pfeile, Gantt. Diese zeichnet rethink selbst als **Formengruppe**. Ein eigener Layout-Kern (`src/engine/`) berechnet die Geometrie.

Dafür bekommt ein Dateneintrag ein zusätzliches `"chart"`-Objekt. Der gleichnamige **Platzhalter** auf der Folie, etwa ein Rechteck namens `Bridge`, wird dann durch das Diagramm ersetzt und gibt Position und Größe vor:

```json
{
  "name": "Bridge",
  "chart": { "type": "waterfall", "datatable": true, "tableRows": [2] },
  "table": [
    [null, {"string":"EBIT 2024"}, {"string":"Volumen"}, {"string":"Kosten"}, {"string":"EBIT 2025"}],
    [{"string":"EBIT"}, {"number":120}, {"number":35}, {"number":-28}, {"string":"e"}],
    [{"string":"Marge"}, {"percentage":8.1}, null, null, {"percentage":8.7}]
  ]
}
```

| Schlüssel in `chart` | Bedeutung (Standard) |
|---|---|
| `type` | `"waterfall"` (weitere Typen siehe [docs/charting-ledger.md](docs/charting-ledger.md)) |
| `datatable` | Datentabelle unter dem Diagramm, Spalten an den Säulen ausgerichtet (`false`) |
| `tableRows` | Zeilennummern aus `table`, die **nur** in der Tabelle stehen (`[]`) |
| `labels` / `categories` / `connectors` | Segmentbeschriftungen, Kategoriebeschriftungen, Wasserfall-Verbinder (`true`) |
| `decimals` | feste Nachkommastellen (automatisch 0–2) |
| `fontSize`, `gap` | Schriftgröße in pt (12), Lücke relativ zur Säulenbreite (0.5) |

**Wasserfall:** Ein `{"string":"e"}` ist ein berechnetes Segment.
- In einer späteren Spalte ist es eine Zwischen- oder Endsumme bis zur Basislinie.
- In der ersten Spalte wird es rückwärts berechnet (absteigende Brücke).

Mehrere Datenzeilen ergeben mehrere Segmente pro Säule, mit höchstens einem `e` je Säule.

**Neu befüllen:** Das Diagramm speichert sein Modell im Tag `RETHINK_MODEL`. Ein späterer Lauf mit derselben Präsentation als Vorlage zeichnet es mit neuen Daten neu. `"chart"`-Angaben im neuen Eintrag überschreiben dabei nur die genannten Optionen.

Welche Diagrammfunktionen schon fertig sind und welche noch offen, zeigt **[docs/charting-ledger.md](docs/charting-ledger.md)**.

## PowerPoint-Add-in (Menüband „rethink“)

Einrichten: `scripts\register.ps1` meldet `bin\rethink-addin.dll` für den aktuellen Benutzer an (nur HKCU). Danach PowerPoint neu starten. Entfernen: `scripts\unregister.ps1`.

| Gruppe / Befehl | Wirkung |
|---|---|
| **Diagramme › Wasserfall einfügen** | Fügt auf der aktuellen Folie einen rethink-Wasserfall mit Beispieldaten und Datentabelle ein und öffnet sofort sein **Datenblatt in Excel**. |
| **Diagramme › Datenblatt öffnen** | Öffnet das Datenblatt des ausgewählten rethink-Diagramms. Ein **Doppelklick** auf das Diagramm tut dasselbe. |
| Erzeugen › Aus .ppttc erzeugen… | Präsentation aus einer `.ppttc`-Datei (wie die Kommandozeile) |
| Elemente › Diagramm einfügen / Elementname setzen | natives Diagramm einfügen bzw. Formen für `.ppttc` benennen |

**Datenblatt-Kreislauf:**
1. Werte, Kategorien und Zeilen in Excel ändern. `e` steht für eine berechnete Summe, Zellfarben werden zu Segmentfarben.
2. Mit **Strg+S speichern**. rethink bemerkt das innerhalb einer Sekunde und zeichnet das Diagramm an derselben Stelle neu.

Das Datenblatt liegt unter `%TEMP%\rethink\`, die Daten selbst stehen im Element (Tag `RETHINK_MODEL`) und reisen mit der Präsentation.

**Kennzeichnung:** Ein ausgewähltes rethink-Diagramm bekommt einen **blauen Leuchtrahmen** mit dem Etikett „rethink · Wasserfall · <Name>“. Das ist ein durchsichtiges Fenster über PowerPoint und verändert das Dokument nicht.

Mit dem Add-in eingefügte und per Kommandozeile erzeugte Diagramme sind dasselbe Format: Beide lassen sich hier bearbeiten und über `.ppttc` neu befüllen.

## Grenzen (bewusst)

- **Eine Vorlage pro `.ppttc`-Datei.** Das Mischen verschiedener Vorlagen (mit unterschiedlichen Mastern) ist für eine spätere Phase vorgesehen.
- **Diagramme in think-cell-Vorlagen** werden nicht gefüllt. Ihre Namen liegen in einem proprietären Datenobjekt, rethink gibt dafür eine Warnung aus. Abhilfe: das Diagramm im Auswahlbereich umbenennen. Die von think-cell gezeichneten Beschriftungen aktualisiert rethink dann aber nicht.
- **Kopierte think-cell-Folien** behalten die think-cell-Formen samt identischer Kennungen und eine eigene Kopie des versteckten Datenobjekts. In PowerPoint öffnen sie sich einwandfrei. Wie sie sich bei aktivem think-cell verhalten, ist **nicht getestet**. Die Kopien sind nicht als weiter bearbeitbare think-cell-Elemente zu betrachten.
- **Punkt- und Blasendiagramme**, Gantt-Diagramme, Harvey Balls, Checkboxen, die Zeile `100%=` und Tabellen werden in Phase 1 nicht unterstützt.
- Text- und Datumswerte in Datenzellen von Diagrammen werden mit einer Warnung ignoriert.
- Die Farben aus `fill` werden nicht in die eingebettete Arbeitsmappe übernommen.
- PowerPoint-Eigenheit: Hat ein Diagramm **genau eine Reihe** und einzeln eingefärbte Punkte, listet die Legende die Kategorien statt des Reihennamens.

## Bauen

Benötigt werden **Visual Studio 2022 Build Tools** (MSVC, CMake) und das **Windows SDK ≥ 10.0.22621**. Weitere Bibliotheken braucht es nicht: rethink nutzt die OPC Packaging API, MSXML 6, `Windows.Data.Json` (C++/WinRT) und URLMon.

```bat
call "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat"
cmake -S . -B %TEMP%\rtk -G "NMake Makefiles" -DCMAKE_BUILD_TYPE=Release
cmake --build %TEMP%\rtk
```

Hinweis: Liegt das Projekt in einem sehr tiefen Pfad, scheitert der Visual-Studio-Generator am 260-Zeichen-Limit von MSBuild. Deshalb wird hier NMake mit einem kurzen Build-Ordner verwendet. Eine fertige Datei liegt unter `bin/rethink.exe`.

## Tests

```powershell
tests\run-tests.ps1 -Exe bin\rethink.exe                     # Validierung und Fehlermeldungen, ohne PowerPoint
samples\make-template.ps1                                    # Testvorlage über PowerPoint erzeugen
bin\rethink.exe build samples\sample.ppttc -o samples\out\sample.pptx
tests\verify.ps1 -Pptx samples\out\sample.pptx -PngDir samples\out\png -ExpectSlides 2 -ExpectCharts 4
tests\roundtrip.ps1 -Pptx samples\out\sample.pptx            # "Daten bearbeiten": Werte bleiben gleich
tests\make-chart-types.ps1                                   # Linie, Fläche, Kreis
bin\rethink.exe build tests\types.ppttc -o tests\out\types-filled.pptx
tests\verify.ps1 -Pptx tests\out\types-filled.pptx -ExpectSlides 1 -ExpectCharts 3
tests\url-test.ps1 -Exe bin\rethink.exe                      # Vorlage per http://localhost
%TEMP%\rtk\engine_test.exe                                   # Layout-Kern + Datenblatt (56 Prüfungen)
tests\make-waterfall-template.ps1; bin\rethink.exe build tests\waterfall.ppttc -o tests\out\waterfall.pptx
tests\waterfall-addin-test.ps1                               # Add-in: Wasserfall, Excel-Datenblatt, Leuchtrahmen (18)
tests\addin-test.ps1                                         # Add-in: Phase-2-Befehle (24)
```

`verify.ps1`, `roundtrip.ps1` und die Vorlagen-Skripte steuern PowerPoint über COM. Wo ein Fenster nötig ist, öffnet sich PowerPoint bzw. Excel kurz minimiert. Läuft PowerPoint bereits, bleibt es geöffnet.

**Getestet (16.–17.10.2026, PowerPoint/Microsoft 365, Windows 11):**
- gestapelte und gruppierte Säulen, Linie mit Markern, Fläche, Kreis
- Leerzeilen-Farbverschiebung und `fill`
- Prozent- und Datumskategorien
- Vervielfältigung auf mehrere Folien
- Textfelder, auch in think-cells Beispielvorlage
- „Daten bearbeiten“ (alle 7 Diagramme unverändert)
- Vorlage per URL
- 8 Validierungsfälle

Balkendiagramme (horizontal) und Ringdiagramme sind nicht eigens getestet.

## Aufbau

| Datei | Aufgabe |
|---|---|
| `src/main.cpp` | Kommandozeile |
| `src/build.cpp` | Ablauf: Vorlage laden, Folien kopieren, Elemente füllen, aufräumen, speichern |
| `src/ppttc.*` | JSON lesen und validieren |
| `src/opc.*` | PPTX/XLSX-Pakete: Teile, Beziehungen, tiefe Kopie, Entfernen verwaister Teile |
| `src/xmldoc.*` | MSXML-DOM mit XPath und Fragment-Import |
| `src/fill_text.*` | Textziele finden und füllen |
| `src/fill_chart.*` | Diagrammreihen, Kategorien, Datenpunktfarben, eingebettete Arbeitsmappe |
| `src/xlsx.*` | XLSX schreiben (Formate, Zellfarben) und von Excel gespeicherte Dateien lesen |
| `src/engine/*` | Layout-Kern gezeichneter Diagramme (ohne COM/XML): Wasserfall-Arithmetik, Geometrie, Beschriftungen, Datentabelle |
| `src/render_element.*` | OOXML-Ausgabe der Grundformen als Gruppe mit `RETHINK_MODEL` |
| `src/datasheet.*` | Datenblatt ↔ Tabelle |
| `addin/elements.*` | Add-in: Grundformen über das PowerPoint-Objektmodell, Einfügen, Datenblatt-Überwachung |
| `addin/events.*`, `addin/overlay.*` | Auswahl- und Doppelklick-Ereignisse, blauer Leuchtrahmen |
