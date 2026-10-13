# Organigramm (HTML + PDF)

Ein Organigramm aus einer Excel-Liste: eine Übersichtsseite plus eine Seite pro Bereich.
Ebenen, direkte Reports und Teamgrößen werden automatisch berechnet.

## Benutzung

1. `vorlage_organigramm.xlsx` öffnen und das Blatt **Daten** ausfüllen (Beispielzeilen sind fiktiv und können gelöscht werden):

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
  (die Excel-Bibliothek liegt in `lib/`, es werden keine externen Skripte geladen).
- Echte Namen müssen daher nie weitergegeben werden. Für Anpassungen am Layout reicht eine anonymisierte
  Struktur (z. B. Namen durch „Person 1“, „Person 2“ … ersetzt).

## Dateien

| Datei | Zweck |
|---|---|
| `organigramm.html` | Organigramm-Seite (Bildschirm + Druck/PDF) |
| `vorlage_organigramm.xlsx` | Excel-Vorlage mit fiktiven Beispieldaten und Anleitung |
| `Organigramm_Mockup.pdf` | Beispiel-PDF (A3 quer) mit den fiktiven Daten |
| `demo-daten.js` | Beispieldaten für die Demo-Ansicht |
| `tools/beispieldaten_erzeugen.py` | erzeugt Vorlage + Demo-Daten neu |
| `tools/pdf_erstellen.mjs` | optional: PDF per Kommandozeile (`node tools/pdf_erstellen.mjs daten.xlsx A3`, benötigt Playwright) |
| `lib/xlsx.full.min.js` | SheetJS 0.18.5 (Apache-2.0), liest Excel-Dateien im Browser |
