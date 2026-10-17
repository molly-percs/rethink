# rethink – Diagramm-Ledger

Diese Liste zeigt, welche Diagrammfunktionen rethink nach dem Vorbild von think-cell abdecken soll und wie weit jede ist.
Grundlage ist das **beschriebene Verhalten** im Benutzerhandbuch von think-cell 14 (Kapitel in Klammern). Es dient nur als Spezifikation, Text und Bilder daraus werden nicht übernommen.
Folienbibliothek und Vorlagen sind bewusst **nicht** Teil dieses Ziels.

**Stand:** ✅ fertig und getestet · 🟡 teilweise · ⬜ offen
**Backend:** **D** = von rethink gezeichnete Formen (Layout-Kern) · **N** = natives PowerPoint-Diagramm (nur `.ppttc`-Vorlagenbefüllung, Phase 1)

> **Ehrlicher Rahmen:** Vollständige Parität mit think-cell ist ein großes Vorhaben. rethink liefert es in senkrechten Scheiben: Jede Scheibe ist ein Diagrammtyp oder eine Funktion, die von den Daten bis zur Folie fertig und getestet ist. Solange in dieser Liste eine Zeile offen ist, ist der Umfang **nicht** vollständig.

## Architektur (Entscheidung)

- **Ein Layout-Kern als einzige Wahrheit** (`src/engine/`): reines C++20 ohne COM und ohne XML. Er übernimmt Datenregeln (`e`, `100%=`, Prozent/absolut, Leerzeilen-Farbsprung), Skalen, Geometrie, Beschriftungstexte, Anmerkungen und die Datentabelle. Ergebnis ist eine Liste von Grundformen (Rechteck, Linie, Pfeil, Text, Polygon) in EMU mit Stilrollen.
- **Dünne Ausgaben** setzen diese Liste um:
  - **OOXML-Ausgabe** (Kommandozeile, `.ppttc`): eine Gruppe aus Formen im Folien-XML.
  - **COM-Ausgabe** (Add-in): dieselben Grundformen über `Shapes.AddShape` / `AddTextbox` / `AddLine`.
  - Beide rechnen **nichts selbst**.
- **Modell am Element:** Die Gruppe trägt das Tag `RETHINK_MODEL`, ein JSON mit Typ, Optionen und Datentabelle. So kann das Add-in sie neu rendern und `.ppttc` sie neu befüllen. Namen und Tags von think-cell werden nie verwendet.
- **Warum gezeichnet statt nativ:** Mekko mit variablen Breiten, Gantt, CAGR- und Differenzpfeile, Summen über Stapeln, Wasserfall-Verbinder und die Datentabelle unter dem Diagramm lassen sich mit nativen Diagrammen nicht abbilden. Native Tabellen können in PowerPoint auch nicht gruppiert werden. Die Grundformen brauchen deshalb eine deterministische Geometrie aus einem einzigen Kern.

## Datenregeln

| Funktion | Verhalten (Spezifikation) | Backend | Stand | Test |
|---|---|---|---|---|
| Datenblatt-Layout | Zeile 0 Kategorien, Spalte 0 Reihennamen | D, N | 🟡 (N ✅) | engine_test, run-tests |
| Leere Zeile | verschiebt das Füllschema um eine Farbe | D, N | 🟡 (N ✅) | engine_test |
| Absolut/Prozent | Prozentwerte werden über die Spaltensumme (oder `100%=`) in Absolutwerte umgerechnet und umgekehrt (element-datasheets) | D | ⬜ | |
| Zeile `100%=` | optionale Summenzeile für Prozentberechnung | D | ⬜ | |
| `e` im Wasserfall | berechnetes Segment; als Zwischensumme bis zur Basislinie; in der ersten Spalte rückwärts gerechnet; höchstens ein `e` pro Säule. **rethink-Konvention in JSON:** `{"string":"e"}` (im Handbuch nicht für JSON festgelegt) | D | ✅ | engine_test (Ende, Mitte, erste Spalte, Nulldurchgang, zwei Segmente, zwei e) |
| Zahlenformat | Beschriftungen mit automatischer Genauigkeit (0–2 Stellen), Dezimaltrennzeichen aus dem Gebietsschema, `decimals` überschreibt | D | 🟡 (kein freies Formatmuster) | engine_test |
| Datumskategorien | Jahr bzw. Monat bei 1.1. bzw. Monatsersten | D, N | 🟡 (N ✅) | |

## Diagrammtypen (quantitative-charts, column-line-area-charts, waterfall, mekko, pie, scatterbubble, gantt)

| Typ | Kernpunkte | Backend | Stand | Test |
|---|---|---|---|---|
| Gestapelte Säulen/Balken | eine Reihe ergibt eine einfache Säule; Balken = gedreht; Abstandsbreite | D (N ✅) | ⬜ | |
| 100 % gestapelt | Summe darf ≠ 100 % sein; Prozent und/oder absolut beschriften; 100 %-Marker | D (N 🟡) | ⬜ | |
| Gruppiert | Segmente nebeneinander | D (N ✅) | ⬜ | |
| **Wasserfall** | aufsteigend und absteigend, `e`, mehrere Segmente pro Säule, Verbinder | D | ✅ (Verbinder nur automatisch, kein manuelles Umhängen) | engine_test, tests/waterfall*.ppttc, verify, PNG |
| Linie / Profil | Interpolation fehlender Werte, Marker, Wertachse bei monoton steigenden X-Werten, zweite Y-Achse | D (N ✅) | ⬜ | |
| Fläche / 100 % Fläche | gestapelt, Summen, Interpolation | D (N ✅) | ⬜ | |
| Kombination | Linien und Säulen je Reihe | D | ⬜ | |
| Mekko % | variable Säulenbreite = Spaltensumme, Y in %, keine Lücken | D | ⬜ | |
| Mekko Einheiten | Zeile „X extent“ für Breiten, absolute Höhen | D | ⬜ | |
| Kreis / Ring | eine Kategorie, Drehung, Innen- und Außenbeschriftung | D (N ✅) | ⬜ | |
| Punkt / Blase | Zeilen = Punkte (X, Y, Größe, Gruppe), Trendlinien, Partitionen | D | ⬜ | |
| Gantt | Kalenderskalen, Aktivitäten, Balken und Meilensteine | D | ⬜ | |

## Anmerkungen, Beschriftungen, Achsen (chartdecorations, textlabels, axes, legends, chart-area-plot-area)

| Funktion | Verhalten | Backend | Stand |
|---|---|---|---|
| Segment-, Serien- und Kategoriebeschriftungen | automatische Platzierung (innen, sonst außen) | D | 🟡 Segment + Kategorie im Wasserfall; Kollisionen zwischen Außenbeschriftungen noch nicht aufgelöst |
| Summen über Säulen | Gesamtwert über dem Stapel | D | ⬜ |
| Wasserfall-Verbinder | waagerechte Linie zwischen benachbarten Säulen | D | ✅ |
| Differenzpfeile | Level- und Gesamtdifferenz; Modus relativ oder absolut | D | ⬜ |
| CAGR-Pfeile | (B/A)^(1/n) − 1, n aus Datumskategorien (30/360) | D | ⬜ |
| Wertlinien | Linie bei Wert oder Mittelwert, Beschriftung links oder rechts | D | ⬜ |
| 100 %-Marker | bei 100 %-Diagrammen | D | ⬜ |
| Serienverbinder | Linien zwischen Segmenten benachbarter Säulen | D | ⬜ |
| Wertachse | Skala, Teilstriche, Gitterlinien | D | ⬜ |
| Kategorie-Abstand | einzelne Lücken vergrößern | D | ⬜ |
| Achsenbruch | Wertebereich unterbrechen | D | ⬜ |
| Legende | Reihen, Reihenfolge | D | ⬜ |
| Segmentreihenfolge | Blatt, umgekehrt, auf- und absteigend | D | ⬜ |
| **Datentabelle unter dem Diagramm** | gruppiert, Spalten an den Säulen ausgerichtet; Zeilen nur in der Tabelle oder auch im Diagramm (charts-with-data-tables) | D | ✅ Wasserfall (`datatable`, `tableRows`); für weitere Typen mit deren Scheibe |

## Bedienung (Add-in)

| Funktion | Stand |
|---|---|
| Tab „rethink“ im Menüband (Phase 2) | ✅ (Phase 2; Migration in den Projektordner läuft) |
| Element einfügen je Diagrammtyp | 🟡 Wasserfall ✅ (COM-Ausgabe aus demselben Layout-Kern) |
| Eingebautes Datenblatt (Tabellen-Look, Symbol unten rechts, Doppelklick) mit Live-Aktualisierung | ✅ Eingabe, Einfügen, Rückgängig, Zeilen/Spalten, %, Zellfarbe (Automation-Test; Maus/Tastatur von Hand ungeprüft) |
| Datenblatt in Excel (optional): live, Schließen ohne Rückfrage | ✅ |
| Kennzeichnung bei Auswahl (Leuchtrahmen + winziges „rethink chart“, ohne Dokumentänderung) | ✅ |
| ▲▼-Griffe je Segment: Wert schrittweise ändern, Verbinder/Summen rechnen mit | ✅ (per Automation; Klick von Hand ungeprüft) |
| Änderungen an Ort und Stelle (eigene Formatierung bleibt) | ✅ |
| Größenänderung per Maus ordnet neu an | ✅ |
| Werte durch Ziehen der Segmentkante ändern | ⬜ |
| `.ppttc` füllt rethink-Elemente (`RETHINK_MODEL`) und Platzhalterformen | ✅ Kommandozeile (Platzhalter → Element, Neuzeichnen mit gespeicherten und überschriebenen Optionen) |

## Scheibenplan

1. ✅ **Wasserfall + Datentabelle** (Layout-Kern, OOXML-Ausgabe, `.ppttc`-Anbindung, Tests)
2. Gestapelt, 100 % und gruppiert inkl. Summen, Wertachse und Legende (gezeichnet)
3. Linie, Fläche, Kombination
4. Mekko (% und Einheiten)
5. Anmerkungen: Differenz, CAGR, Wertlinie, 100 %-Marker, Serienverbinder
6. Kreis/Ring, Punkt/Blase
7. Gantt
8. Add-in: Einfügen je Typ, Datenblatt-Kreislauf, COM-Ausgabe
