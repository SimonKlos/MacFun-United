# Organigramm (HTML + PDF)

`organigramm.html` ist **eine einzige, eigenständige Datei**: herunterladen, per Doppelklick im Browser öffnen,
Excel laden, als PDF drucken. Keine Installation, kein Internet nötig, die Daten bleiben lokal im Browser.

## Was erzeugt wird (je Org-Einheit 1, z. B. „Männerfußball“)

1. **Organigramm in FTE**: schwarz = gesamt, dunkelgrau = Bereich (N-1, summiert), hellgrau = Org-Einheit,
   weiße Kästchen = Personen mit FTE. Praktikant/innen & Aushilfen stehen in **blauer** Schrift.
2. **Organigramm in Köpfen**: gleiche Struktur, gezählt nach Personen.
3. **Aufbau nach Bereichen**: Wasserfall je Bereich (N-1) bis Gesamt, links in FTE, rechts in Köpfen.

## Erwartete Excel-Spalten

`N-1 (Person) | N-2 (Person) | N-3 (Person) | Org-Einheit 1 | Org-Einheit 2 | Org-Einheit 3 | Mitarbeitername | Summe von FTE | Personalfall | Organigramm final`

| Spalte | Verwendung |
|---|---|
| Org-Einheit 1 | oberste Einheit / Seitentitel (Teil vor „ / “, z. B. „Männerfußball“) |
| Organigramm final | Bereich auf Ebene N-1 (dunkelgraue Kästen, Wasserfall-Zeilen) |
| Org-Einheit 2 | Org-Einheit darunter (hellgraue Kästen), Präfix „Männerfußball - “ wird entfernt; umschaltbar auf Org-Einheit 3 |
| Mitarbeitername | Personen-Kästchen; jede Zeile = 1 Kopf |
| Summe von FTE | FTE je Person (Komma oder Punkt) |
| Org-Einheit 2/3 enthält „Prakti“ oder „Aushilf“ | Person wird blau dargestellt und direkt unter dem Bereich gezeigt |
| N-1/N-2/N-3 (Person), Personalfall | werden aktuell nicht verwendet |

Die Kopfzeile darf auch weiter unten stehen (z. B. Pivot-Tabelle mit Filtern darüber). Summenzeilen
(„Gesamtergebnis“) und Zeilen ohne Mitarbeitername werden ignoriert. Leeres „Org-Einheit 1“ wird von oben übernommen.

## Dateien

| Datei | Zweck |
|---|---|
| `organigramm.html` | Organigramm-Seite (Bildschirm + Druck/PDF), alles in einer Datei |
| `testdaten_organigramm.xlsx` | fiktive Testdaten im Aufbau der Original-Excel (64,75 FTE / 74 Köpfe) |
| `Organigramm_Mockup.pdf` | Beispiel-PDF (A3 quer) mit den Testdaten |
| `tools/beispieldaten_erzeugen.py` | erzeugt die Testdaten neu und setzt sie als Demo in die HTML ein |
| `tools/pdf_erstellen.mjs` | optional: PDF per Kommandozeile (`node tools/pdf_erstellen.mjs daten.xlsx A3`, benötigt Playwright) |
| `LICENSE-sheetjs.txt` | Lizenz der eingebetteten Excel-Bibliothek SheetJS 0.18.5 (Apache-2.0) |
