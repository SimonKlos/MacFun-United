# Organigramm-Simulator

Erstellt aus dem **HR-Works-Mitarbeiterauszug** (Excel) Organigramme als Bildschirmansicht und PDF.
`organigramm.html` ist **eine einzige, eigenständige Datei**: per Doppelklick im Browser öffnen, Excel laden,
Optionen wählen, als PDF drucken. Keine Installation, kein Internet nötig, die Daten bleiben lokal im Browser.

Ausführliche Bedienung: **`Anleitung_Organigramm-Simulator.pdf`**

## Ausgaben (einzeln an-/abwählbar)

| Seite | Inhalt |
|---|---|
| Organigramm in FTE | Analytics: direkt unterstellte FTE (rote Zahl), Team-FTE, FTE je Person |
| Organigramm in Köpfen | Analytics: direkt unterstellte Personen, Teamgröße |
| Organigramm ohne Zahlen | Kommunikationsversion, nur Namen/Positionen |
| Aufbau je Bereich | Wasserfall FTE + Köpfe je Bereich und Kennzahlentabelle (Festangestellte, Aushilfen, Ø FTE/Kopf) |
| Detailseite je Bereich | je Bereich eine eigene Seite mit allen Ebenen |

Weitere Optionen: Gesellschaft, Bereich, Stichtag, Tiefe der Übersicht, Namen (voll / Initialen / anonym),
Position ein/aus, Aushilfen (hervorheben / zusammenfassen / ausblenden), FTE-Basis (Standard 40 Std.), A3/A4.

## Verwendete Spalten

`Personalnummer` und `Vorgesetzter-Personalnummer` (Hierarchie, Pflicht) · `Vorname`, `Name` · `Position` ·
`Wochenarbeitsstunden` (FTE = Std. ÷ 40) · `Hierarchiefeld` (Aushilfe/Werkstudent/Praktikum) · `Arbeitsverhältnis` (Ersatz) ·
`Gesellschaft / Mandant` (oben rechts, Filter) · `Orga-Einheit, Level 1` (Bereichsfilter) · `Name der Organisationseinheit` (Bereichsnamen) ·
`Eintrittsdatum` / `Austrittsdatum` (Stichtag). Alle anderen Spalten werden ignoriert.

## Dateien

| Datei | Zweck |
|---|---|
| `organigramm.html` | der Simulator, alles in einer Datei |
| `Anleitung_Organigramm-Simulator.pdf` | Anleitung für die Personalabteilung |
| `testdaten_gross.xlsx`, `testdaten_klein.xlsx` | fiktive Testdaten mit exakt den Spalten des HR-Works-Auszugs (133 bzw. 25 aktive Personen); die kleine Einheit ist im Simulator vorausgefüllt |
| `Beispiel_gross.pdf`, `Beispiel_klein.pdf` | Beispiel-Ausdrucke mit den Testdaten |
| `tools/beispieldaten_erzeugen.py` | erzeugt die Testdaten neu und setzt sie als Demo in die HTML ein |
| `tools/anleitung.html` | Quelle der Anleitung |
| `tools/pdf_erstellen.mjs` | optional: PDF per Kommandozeile (`node tools/pdf_erstellen.mjs daten.xlsx A3`, benötigt Playwright) |
| `LICENSE-sheetjs.txt` | Lizenz der eingebetteten Excel-Bibliothek SheetJS 0.18.5 (Apache-2.0) |

Echte HR-Auszüge gehören **nicht** in dieses Repository.
