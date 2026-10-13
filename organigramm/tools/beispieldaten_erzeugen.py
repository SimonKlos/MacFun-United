"""Erzeugt fiktive Beispieldaten fuer das Organigramm-Mockup.

Ausgabe:
  ../vorlage_organigramm.xlsx   Excel-Vorlage (Blatt "Daten" + Blatt "Anleitung")
  ../organigramm.html           dieselben Daten werden dort als Demo eingesetzt

Alle Namen sind frei erfunden. Aufruf:  python3 beispieldaten_erzeugen.py
"""
import itertools
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

OUT = Path(__file__).resolve().parent.parent

VORNAMEN = ["Anna", "Ben", "Clara", "David", "Elif", "Felix", "Greta", "Hannes", "Ida", "Jonas",
            "Katrin", "Lukas", "Mara", "Niklas", "Olga", "Paul", "Rana", "Stefan", "Tina", "Uwe",
            "Vera", "Wim", "Yasmin", "Zoe", "Moritz", "Lea", "Tobias", "Nora", "Emil", "Sophie"]
NACHNAMEN = ["Albers", "Brandt", "Cordes", "Dietz", "Engel", "Fuchs", "Graf", "Hahn", "Imhof",
             "Jansen", "Kraus", "Lorenz", "Maier", "Nowak", "Ott", "Peters", "Quast", "Roth",
             "Seidel", "Thiel", "Ulrich", "Vogt", "Weber", "Yilmaz", "Ziegler", "Berger", "Kaiser",
             "Lang", "Möller", "Schulte", "Hoffmann", "Busch", "Arnold", "Kuhn", "Winter"]

_namen = (f"{v} {n}" for n, v in zip(itertools.cycle(NACHNAMEN),
                                      itertools.islice(itertools.cycle(VORNAMEN), 400)))
_seen = set()


def name():
    for n in _namen:
        if n not in _seen:
            _seen.add(n)
            return n
    raise RuntimeError("Namen aufgebraucht")


rows = []  # (ID, Name, Funktion, Bereich, Vorgesetzter)
_ids = itertools.count(1)


def add(funktion, bereich, chef):
    pid = f"P{next(_ids):03d}"
    rows.append((pid, name(), funktion, bereich, chef or ""))
    return pid


def team(chef, bereich, funktion, n):
    return [add(funktion, bereich, chef) for _ in range(n)]


ceo = add("Geschäftsführung", "Geschäftsführung", None)

# 1) Vertrieb - 6 Ebenen
b = "Vertrieb"
bl = add("Bereichsleitung Vertrieb", b, ceo)
for region in ["Nord", "Süd"]:
    rl = add(f"Regionalleitung {region}", b, bl)
    for gebiet in (["Hamburg", "Berlin"] if region == "Nord" else ["München"]):
        gl = add(f"Gebietsleitung {gebiet}", b, rl)
        for t in range(2 if gebiet != "Berlin" else 1):
            tl = add(f"Teamleitung {gebiet} {t + 1}", b, gl)
            for k in range(2 if t == 0 else 1):
                kam = add("Key Account Manager", b, tl)
                team(kam, b, "Sales Representative", 2 + (k % 2))
team(bl, b, "Vertriebsassistenz", 1)

# 2) Finanzen - 3 Ebenen
b = "Finanzen"
bl = add("Bereichsleitung Finanzen", b, ceo)
for abt, n in [("Buchhaltung", 4), ("Controlling", 3), ("Treasury", 2)]:
    al = add(f"Leitung {abt}", b, bl)
    team(al, b, f"Referent/in {abt}", n)

# 3) IT - 4 Ebenen
b = "IT"
bl = add("CIO / Bereichsleitung IT", b, ceo)
for abt, teams in [("Infrastruktur", [("Netzwerk", 3), ("Service Desk", 4)]),
                   ("Software", [("Backend", 3), ("Frontend", 2), ("Daten", 2)])]:
    al = add(f"Leitung {abt}", b, bl)
    for tname, n in teams:
        tl = add(f"Teamleitung {tname}", b, al)
        team(tl, b, f"Engineer {tname}", n)
add("IT-Sicherheitsbeauftragte/r", b, bl)

# 4) Personal - 2 Ebenen
b = "Personal"
bl = add("Bereichsleitung Personal", b, ceo)
for f in ["HR Business Partner", "HR Business Partner", "Recruiting", "Lohn & Gehalt", "Personalentwicklung"]:
    add(f, b, bl)

# 5) Produktion - 5 Ebenen
b = "Produktion"
bl = add("Bereichsleitung Produktion", b, ceo)
for werk in ["Werk A", "Werk B"]:
    wl = add(f"Werksleitung {werk}", b, bl)
    for schicht in (["Früh", "Spät"] if werk == "Werk A" else ["Tag"]):
        sl = add(f"Schichtleitung {schicht} ({werk})", b, wl)
        for linie in range(2):
            ml = add(f"Meister/in Linie {linie + 1}", b, sl)
            team(ml, b, "Fachkraft Produktion", 3 if linie == 0 else 2)
    add(f"Qualitätssicherung {werk}", b, wl)

HEADER = ["ID", "Name", "Funktion", "Bereich", "Vorgesetzter"]

# ---------- Excel ----------
wb = Workbook()
ws = wb.active
ws.title = "Daten"
ws.append(HEADER)
for r in rows:
    ws.append(list(r))

head_fill = PatternFill("solid", fgColor="C85A5A")
for c in ws[1]:
    c.font = Font(bold=True, color="FFFFFF")
    c.fill = head_fill
    c.alignment = Alignment(vertical="center")
ws.row_dimensions[1].height = 22
for col, w in zip("ABCDE", [9, 24, 36, 20, 16]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:E{len(rows) + 1}"
zebra = PatternFill("solid", fgColor="F7F3F3")
for i in range(2, len(rows) + 2, 2):
    for c in ws[i]:
        c.fill = zebra

dv = DataValidation(type="list", formula1='"Geschäftsführung,Vertrieb,Finanzen,IT,Personal,Produktion"',
                    allow_blank=True, showErrorMessage=False)
ws.add_data_validation(dv)
dv.add(f"D2:D1000")

an = wb.create_sheet("Anleitung")
anleitung = [
    ("So füllen Sie die Vorlage aus", True),
    ("", False),
    ("1. Eine Zeile pro Person im Blatt „Daten“. Die Beispielzeilen sind frei erfunden und können gelöscht werden.", False),
    ("2. ID: beliebiges eindeutiges Kürzel (z. B. P001 oder Personalnummer). Pflichtfeld.", False),
    ("3. Name, Funktion: erscheinen auf der Karte im Organigramm.", False),
    ("4. Bereich: Name des Bereichs. Maßgeblich ist der Eintrag bei der Bereichsleitung (Seitentitel).", False),
    ("5. Vorgesetzter: ID (empfohlen) ODER exakter Name der direkten Führungskraft.", False),
    ("   Genau eine Person (Geschäftsführung) bleibt hier leer – sie ist die Spitze.", False),
    ("6. Ebenen und Teamgrößen werden automatisch berechnet – nichts zählen, nichts formatieren.", False),
    ("7. Reihenfolge: Personen erscheinen in der Reihenfolge der Zeilen.", False),
    ("", False),
    ("Datei danach in organigramm.html über „Excel laden“ öffnen. Die Daten werden nur im Browser", False),
    ("verarbeitet, nicht hochgeladen und nicht gespeichert.", False),
]
for text, bold in anleitung:
    an.append([text])
    if bold:
        an.cell(an.max_row, 1).font = Font(bold=True, size=14, color="9E3B3B")
an.column_dimensions["A"].width = 110

wb.save(OUT / "vorlage_organigramm.xlsx")

# ---------- Demo-Daten in organigramm.html einsetzen ----------
data = [dict(zip(HEADER, r)) for r in rows]
html_path = OUT / "organigramm.html"
html = html_path.read_text(encoding="utf-8")
start, end = "/*DEMO-DATEN-START*/", "/*DEMO-DATEN-ENDE*/"
a, b = html.index(start) + len(start), html.index(end)
html_path.write_text(html[:a] + json.dumps(data, ensure_ascii=False) + html[b:], encoding="utf-8")

print(f"{len(rows)} Personen geschrieben.")
