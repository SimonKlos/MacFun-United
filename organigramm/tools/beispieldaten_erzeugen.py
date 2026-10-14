"""Erzeugt fiktive Testdaten im Aufbau der Original-Excel.

Ausgabe:
  ../testdaten_organigramm.xlsx   Testdatei (gleiche Spalten wie die Original-Excel)
  ../organigramm.html             dieselben Daten werden dort als Demo eingesetzt

Alle Namen sind frei erfunden. Aufruf:  python3 beispieldaten_erzeugen.py
"""
import itertools
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

OUT = Path(__file__).resolve().parent.parent

VORNAMEN = ["Anna", "Ben", "Clara", "David", "Elif", "Felix", "Greta", "Hannes", "Ida", "Jonas",
            "Katrin", "Lukas", "Mara", "Niklas", "Olga", "Paul", "Rana", "Stefan", "Tina", "Uwe",
            "Vera", "Wim", "Yasmin", "Zoe", "Moritz", "Lea", "Tobias", "Nora", "Emil", "Sophie",
            "Jan", "Mia", "Tim", "Sara", "Leon", "Julia", "Max", "Hanna", "Erik", "Laura", "Kai"]
NACHNAMEN = ["Albers", "Brandt", "Cordes", "Dietz", "Engel", "Fuchs", "Graf", "Hahn", "Imhof",
             "Jansen", "Kraus", "Lorenz", "Maier", "Nowak", "Ott", "Peters", "Quast", "Roth",
             "Seidel", "Thiel", "Ulrich", "Vogt", "Weber", "Yilmaz", "Ziegler", "Berger", "Kaiser",
             "Lang", "Möller", "Schulte", "Hoffmann", "Busch", "Arnold", "Kuhn", "Winter", "Becker"]
_namen = (f"{v} {n}" for v, n in zip(itertools.cycle(VORNAMEN), itertools.cycle(NACHNAMEN)))

ORG1 = "Männerfußball / Men's Football"
PRAKTI = "Männerfußball - Praktis & Aushilfen"

# Organigramm final -> (Leitung N-1, [(Org-Einheit-2-Kurzname | None = direkt, [FTE je Person])])
STRUKTUR = {
    "Medizinischer Bereich": [
        ("Physiotherapie", [1, 1, 1, 1]),
        ("Ernährung", [1, 1, 1, 1]),
        ("Rehabilitation", [1, 1]),
        ("Wissenschaftl. Abteilung", [1, 1, 0.75]),
        ("Mannschaftsärzte", [0.25, 0.25]),
        (PRAKTI, [0.25, 0.25]),
    ],
    "Sportliche Leitung": [
        ("Sportliche Leitung", [1, 0.5]),
        ("Trainerteam", [1] * 10 + [0.75, 0.5]),
        ("Lizenzspieler", [1] * 7 + [0.75]),
        ("Analyse & Technologie", [1, 1]),
        ("Video & Spielanalyse", [1, 1, 1, 1]),
        ("Scouting", [1] * 17 + [0.75, 0.5]),
        (PRAKTI, [0.25, 0.25]),
    ],
    "Referent Sport-VS": [
        ("Referent Sport-VS", [1, 1]),
    ],
    "EFI": [
        ("EFI", [1, 1, 1]),
        (PRAKTI, [0.25, 0.25, 0.25]),
    ],
}

HEADER = ["N-1 (Person)", "N-2 (Person)", "N-3 (Person)", "Org-Einheit 1", "Org-Einheit 2",
          "Org-Einheit 3", "Mitarbeitername", "Summe von FTE", "Personalfall", "Organigramm final"]

rows = []
for bereich, einheiten in STRUKTUR.items():
    n1 = next(_namen)
    for einheit, ftes in einheiten:
        org2 = einheit if einheit == PRAKTI else f"Männerfußball - {einheit}"
        n2 = "" if einheit in (PRAKTI, bereich) else next(_namen)
        for fte in ftes:
            rows.append([n1, n2, "", ORG1, org2, "", next(_namen), fte, "", bereich])

# ---------- Excel ----------
wb = Workbook()
ws = wb.active
ws.title = "Daten"
ws.append(HEADER)
for r in rows:
    ws.append(r)
head_fill = PatternFill("solid", fgColor="BDD7EE")
thin = Side(style="thin", color="9BC2E6")
for c in ws[1]:
    c.font = Font(bold=True)
    c.fill = head_fill
    c.border = Border(bottom=thin)
    c.alignment = Alignment(vertical="center")
for col, w in zip("ABCDEFGHIJ", [18, 18, 18, 30, 36, 20, 22, 14, 13, 24]):
    ws.column_dimensions[col].width = w
for row in ws.iter_rows(min_row=2, min_col=8, max_col=8):
    row[0].number_format = "0.00"
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:J{len(rows) + 1}"
wb.save(OUT / "testdaten_organigramm.xlsx")

# ---------- Demo-Daten in organigramm.html einsetzen ----------
data = [dict(zip(HEADER, r)) for r in rows]
html_path = OUT / "organigramm.html"
html = html_path.read_text(encoding="utf-8")
start, end = "/*DEMO-DATEN-START*/", "/*DEMO-DATEN-ENDE*/"
a, b = html.index(start) + len(start), html.index(end)
html_path.write_text(html[:a] + json.dumps(data, ensure_ascii=False) + html[b:], encoding="utf-8")

print(f"{len(rows)} Personen, {sum(r[7] for r in rows):g} FTE")
