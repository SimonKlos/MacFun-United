# Organigramm (HTML + PDF)

Ein Organigramm aus einer Excel-Liste: eine Übersichtsseite mit allen Ebenen plus eine Seite pro Bereich.

**`organigramm.html` ist eine einzige, eigenständige Datei** – einfach herunterladen, per Doppelklick
öffnen, fertig. Keine Installation, kein Internet nötig.
Ebenen, direkte Reports und Teamgrößen werden automatisch berechnet.

## Benutzung

1. In der Seite auf **Vorlage herunterladen** klicken (oder `vorlage_organigramm.xlsx` nehmen) und und das Blatt **Daten** ausfüllen (Beispielzeilen sind fiktiv und können gelöscht werden):

   | ID | Name | Funktion | Bereich | Vorgesetzter |
   |----|------|----------|---------|--------------|
   | P001 | … | Geschäftsführung | Geschäftsführung | *(leer = Spitze)* |
   | P002 | … | Bereichsleitung Vertrieb | Vertrieb | P001 |

   *Vorgesetzter* darf die ID **oder** der exakte Name sein.
2. `organigramm.html` im Browser öffnen → **Excel laden** (oder die Datei ins Fenster ziehen).
3. **Drucken / PDF** → im Druckdialog „Als PDF speichern“. Format (A3/A4 quer) oben rechts wählen.
   Jede Seite wird automatisch passend skaliert; sehr tiefe/breite Bereiche werden automatisch kompakt (eingerückt) dargestellt.

## Datenschutz

- Die Excel-Datei wird **nur lokal im Browser** gelesen: kein Upload, keine Speicherung, kein Netzwerkzugriff
  (die Excel-Bibliothek ist in die HTML-Datei eingebettet, es werden keine externen Skripte geladen).
- Echte Namen müssen daher nie weitergegeben werden. Für Anpassungen am Layout reicht eine anonymisierte
  Struktur (z. B. Namen durch „Person 1“, „Person 2“ … ersetzt).

## Dateien

| Datei | Zweck |
|---|---|
| `organigramm.html` | Organigramm-Seite (Bildschirm + Druck/PDF), alles in einer Datei |
| `vorlage_organigramm.xlsx` | Excel-Vorlage mit fiktiven Beispieldaten und Anleitung |
| `Organigramm_Mockup.pdf` | Beispiel-PDF (A3 quer) mit den fiktiven Daten |
| `tools/beispieldaten_erzeugen.py` | erzeugt Vorlage neu und setzt die Demo-Daten in die HTML ein |
| `tools/pdf_erstellen.mjs` | optional: PDF per Kommandozeile (`node tools/pdf_erstellen.mjs daten.xlsx A3`, benötigt Playwright) |
| `LICENSE-sheetjs.txt` | Lizenz der eingebetteten Excel-Bibliothek SheetJS 0.18.5 (Apache-2.0) |
