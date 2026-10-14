"""Erzeugt fiktive Testdaten im Spaltenaufbau des HR-Works-Auszugs.

Ausgabe:
  ../testdaten_gross.xlsx   größere Einheit (ca. 140 Personen, bis 5 Ebenen)
  ../testdaten_klein.xlsx   kleinere Einheit (ca. 25 Personen)
  ../organigramm.html       die kleine Einheit wird dort als Beispiel vorausgefüllt

Alle Namen und Werte sind frei erfunden. Aufruf:  python3 beispieldaten_erzeugen.py
"""
import datetime as dt
import itertools
import json
import zlib
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

OUT = Path(__file__).resolve().parent.parent

# Spalten exakt wie im HR-Works-Auszug (A..AL)
HEADER = ["Benutzerkennung", "Name", "Vorname", "Personalnummer", "Vorgesetzter-Personalnummer",
          "Organisationseinheiten-Schlüssel", "Orga-Einheit, Level 1", "Orga-Einheit Vorstandsressort",
          "Position", "Orga-Ebene der Position", "Titel laut Vertrag", "Gesellschaft / Mandant",
          "Sport / Nicht-Sport", "Hierarchiefeld", "Sport: Start aktueller Vertrag", "Arbeitsverhältnis",
          "Status (NLZ)", "Geburtsdatum", "Eintrittsdatum", "Austrittsdatum", "Sport: Start aktueller Vertrag",
          "E-Mail", "DATEV-Personalnummer", "Berechtigung - Darf Homeoffice beantragen",
          "Berechtigung - Kinderdaten bearbeiten", "Geschlecht", "Arbeitsplan", "Wochenarbeitsstunden",
          "Hat Dienstwagenanspruch", "Anrede", "Name der Organisationseinheit", "Betriebsstätte-Schlüssel",
          "Hat Dienstwagen", "Urlaubsart-Schlüssel", "Wochenarbeitstage", "Zeiterfassungsart",
          "Zeiterfassungsart-Schlüssel", "Stundenlohn"]

HF = {  # Hierarchiefeld
    "vorstand": "Vorstand / Executive Board",
    "dir": "Direktor:in / Director",
    "al": "Abteilungsleitung / Head of",
    "tl": "Teamleitung / Team Lead",
    "ma": "Mitarbeiter:in / Employee",
    "aus": "Aushilfe & Werkstudent:in / Temp & Working Student",
    "prak": "Praktikant:in / Intern",
}
AV = {"aus": "Geringfügig beschäftigt", "prak": "Praktikum"}
TEMP = ("aus", "prak")

VORNAMEN = ["Anna", "Ben", "Clara", "David", "Elif", "Felix", "Greta", "Hannes", "Ida", "Jonas",
            "Katrin", "Lukas", "Mara", "Niklas", "Olga", "Paul", "Rana", "Stefan", "Tina", "Uwe",
            "Vera", "Wim", "Yasmin", "Zoe", "Moritz", "Lea", "Tobias", "Nora", "Emil", "Sophie",
            "Jan", "Mia", "Tim", "Sara", "Leon", "Julia", "Max", "Hanna", "Erik", "Laura", "Kai"]
NACHNAMEN = ["Albers", "Brandt", "Cordes", "Dietz", "Engel", "Fuchs", "Graf", "Hahn", "Imhof",
             "Jansen", "Kraus", "Lorenz", "Maier", "Nowak", "Ott", "Peters", "Quast", "Roth",
             "Seidel", "Thiel", "Ulrich", "Vogt", "Weber", "Yilmaz", "Ziegler", "Berger", "Kaiser",
             "Lang", "Möller", "Schulte", "Hoffmann", "Busch", "Arnold", "Kuhn", "Winter", "Becker"]


def P(position, hf="ma", hours=40, ae=None, kids=(), area=None, **extra):
    """Knoten im Testbaum: Position, Hierarchiefeld-Kürzel, Wochenstunden, Org-Einheit, Unterstellte."""
    return dict(position=position, hf=hf, hours=hours, ae=ae, kids=list(kids), area=area, **extra)


def team(n, position, hf="ma", hours=40, **kw):
    hrs = hours if isinstance(hours, list) else [hours] * n
    return [P(position, hf, hrs[i], **kw) for i in range(n)]


def build(tree, gesellschaft, prefix, name_offset, root_self_ref):
    names = zip(itertools.islice(itertools.cycle(VORNAMEN), name_offset, None),
                itertools.islice(itertools.cycle(NACHNAMEN), name_offset, None))
    seq = itertools.count(1)
    rows = []

    def walk(node, boss, area, ae, depth):
        area = node["area"] or area
        ae = node["ae"] or ae
        vor, nach = next(names)
        n = next(seq)
        pnr = f"{prefix}{10000 + n}"
        hf = node["hf"]
        hours = node["hours"]
        av = AV.get(hf, "Vollzeit" if hours >= 40 else "Teilzeit")
        eintritt = node.get("eintritt", dt.datetime(2014 + n % 11, 1 + n % 12, 1))
        r = dict.fromkeys(HEADER, "")
        r.update({
            "Benutzerkennung": 5000 + n, "Name": nach, "Vorname": vor, "Personalnummer": pnr,
            "Vorgesetzter-Personalnummer": boss or (pnr if root_self_ref else ""),
            "Organisationseinheiten-Schlüssel": 1000 + zlib.crc32(ae.encode()) % 900,
            "Orga-Einheit, Level 1": area, "Orga-Einheit Vorstandsressort": "Ressort Vorstand",
            "Position": node["position"], "Orga-Ebene der Position": f"E{depth}",
            "Titel laut Vertrag": "Aushilfe" if hf == "aus" else node["position"],
            "Gesellschaft / Mandant": gesellschaft,
            "Sport / Nicht-Sport": "Sport" if area in ("Sport", "Nachwuchsleistungszentrum") else "Nicht-Sport",
            "Hierarchiefeld": HF[hf], "Arbeitsverhältnis": av, "Status (NLZ)": "n/a",
            "Eintrittsdatum": eintritt, "Austrittsdatum": node.get("austritt", ""),
            "Berechtigung - Darf Homeoffice beantragen": "Nein" if hf in TEMP else "Ja",
            "Berechtigung - Kinderdaten bearbeiten": "Nein", "Arbeitsplan": f"{hours}h (Mo-Fr)",
            "Wochenarbeitsstunden": hours, "Hat Dienstwagenanspruch": "Nein",
            "Name der Organisationseinheit": ae, "Betriebsstätte-Schlüssel": "Standort Mitte",
            "Hat Dienstwagen": "Nein", "Urlaubsart-Schlüssel": "30", "Stundenlohn": "n/a",
            "Zeiterfassungsart": "Zeiterfassung Standard", "Zeiterfassungsart-Schlüssel": 1,
        })
        rows.append(r)
        for k in node["kids"]:
            walk(k, pnr, area, ae, depth + 1)

    walk(tree, None, tree["area"], tree["ae"], 0)
    return rows


# ---------------- größere Einheit ----------------
GROSS = P("Vorstandsvorsitz", "vorstand", area="Vorstand", ae="Vorstand", kids=[
    P("Assistenz des Vorstands", hours=30),
    P("Referent Strategie"),
    P("Direktor Sport", "dir", area="Sport", ae="Sport", kids=[
        P("Teammanager"),
        P("Cheftrainer", "al", ae="Trainerteam", kids=team(3, "Co-Trainer") + [P("Torwarttrainer")] + team(2, "Athletiktrainer")),
        P("Leitung Medizin", "al", ae="Medizinische Abteilung", kids=team(4, "Physiotherapeut/in") + [P("Mannschaftsarzt", hours=20)] + team(2, "Rehatrainer/in", hours=[40, 30])),
        P("Leitung Scouting", "al", ae="Scouting", kids=team(6, "Scout") + team(2, "Werkstudent/in Scouting", "aus", 20)),
        P("Leitung Spielanalyse", "al", ae="Spielanalyse", kids=team(3, "Videoanalyst/in") + [P("Werkstudent/in Analyse", "aus", 20)]),
    ]),
    P("Direktor Nachwuchsleistungszentrum", "dir", area="Nachwuchsleistungszentrum", ae="NLZ", kids=[
        P("Leitung Leistungsbereich", "al", ae="NLZ Leistungsbereich", kids=[
            P("Cheftrainer U19", "tl", kids=team(2, "Co-Trainer U19") + [P("Aushilfstrainer U19", "aus", 10)]),
            P("Cheftrainer U17", "tl", kids=team(2, "Co-Trainer U17") + [P("Aushilfstrainer U17", "aus", 10)]),
        ]),
        P("Leitung Grundlagenbereich", "al", ae="NLZ Grundlagenbereich", kids=[
            P("Trainer U15", "tl", kids=[P("Co-Trainer U15", hours=30)] + team(2, "Aushilfstrainer U15", "aus", 10)),
            P("Trainer U14", "tl", hours=30, kids=team(2, "Aushilfstrainer U14", "aus", 8)),
            P("Trainer U13", "tl", hours=30, kids=team(2, "Aushilfstrainer U13", "aus", 8)),
        ]),
        P("Pädagogische Leitung", hours=30, kids=[P("Internatsbetreuer/in")] + team(2, "Aushilfe Internat", "aus", 12)),
    ]),
    P("Direktor Marketing & Vertrieb", "dir", area="Marketing & Vertrieb", ae="Marketing & Vertrieb", kids=[
        P("Leitung Marketing", "al", ae="Marketing", kids=[
            P("Teamleitung Social Media", "tl", kids=team(3, "Social Media Manager/in") + team(2, "Werkstudent/in Social Media", "aus", 20)),
            P("Teamleitung Content", "tl", kids=team(3, "Content Creator", hours=[40, 40, 30]) + [P("Praktikant/in Content", "prak", 40)]),
        ]),
        P("Leitung Sponsoring", "al", ae="Sponsoring", kids=team(4, "Key Account Manager/in")),
        P("Leitung Ticketing", "al", ae="Ticketing", kids=[
            P("Teamleitung Service Center", "tl", kids=team(2, "Servicemitarbeiter/in") + team(6, "Aushilfe Service Center", "aus", 12)),
        ]),
    ]),
    P("Direktor Finanzen & Verwaltung", "dir", area="Finanzen & Verwaltung", ae="Finanzen & Verwaltung", kids=[
        P("Leitung Buchhaltung", "al", ae="Buchhaltung", kids=team(4, "Buchhalter/in", hours=[40, 40, 30, 20])),
        P("Leitung Controlling", "al", ae="Controlling", kids=team(3, "Controller/in") + [P("Werkstudent/in Controlling", "aus", 20)]),
        P("Leitung Personal", "al", ae="Personal", kids=team(3, "HR Business Partner/in", hours=[40, 40, 32]) + [P("Payroll Specialist", eintritt=dt.datetime(2026, 11, 1))]),
        P("Syndikusrechtsanwält/in", hours=40, ae="Recht", kids=[P("Paralegal", hours=30, austritt=dt.datetime(2026, 6, 30))]),
    ]),
    P("Leitung Merchandising", "al", area="Merchandising", ae="Merchandising", kids=[
        P("Teamleitung Fanshop", "tl", ae="Fanshop", kids=team(3, "Verkäufer/in") + team(8, "Aushilfe Fanshop", "aus", [10, 10, 12, 12, 15, 15, 8, 8])),
        P("Teamleitung E-Commerce", "tl", ae="E-Commerce", kids=team(3, "E-Commerce Manager/in")),
        P("Teamleitung Logistik", "tl", ae="Lager", kids=team(2, "Lagerist/in") + team(3, "Aushilfe Lager", "aus", 15)),
    ]),
    P("Leitung IT", "al", area="IT", ae="IT", kids=team(3, "IT-Administrator/in") + team(2, "Softwareentwickler/in") + [P("Werkstudent/in IT", "aus", 20)]),
])

# ---------------- kleinere Einheit ----------------
KLEIN = P("Geschäftsführung", "vorstand", area="Geschäftsführung", ae="Geschäftsführung", kids=[
    P("Buchhaltung", hours=20, ae="Verwaltung"),
    P("Leitung Events", "al", area="Events", ae="Events", kids=team(3, "Eventmanager/in") + team(4, "Aushilfe Events", "aus", 10)),
    P("Leitung Gastronomie", "al", area="Gastronomie", ae="Gastronomie", kids=[
        P("Teamleitung Küche", "tl", ae="Küche", kids=team(3, "Koch/Köchin", hours=[40, 40, 30]) + team(2, "Aushilfe Küche", "aus", 12)),
        P("Teamleitung Service", "tl", ae="Service", kids=team(2, "Servicekraft") + team(5, "Aushilfe Service", "aus", [8, 10, 10, 12, 15])),
    ]),
])

DATASETS = {
    "gross": ("testdaten_gross.xlsx", build(GROSS, "Muster Sport AG", "A", 0, True)),
    "klein": ("testdaten_klein.xlsx", build(KLEIN, "Muster Service GmbH", "S", 17, False)),
}

head_fill = PatternFill("solid", fgColor="4472C4")
for key, (fname, rows) in DATASETS.items():
    wb = Workbook()
    ws = wb.active
    ws.title = "Tabelle1"
    ws.append(HEADER)
    for r in rows:
        ws.append([r[h] for h in HEADER])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head_fill
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for i in range(1, len(HEADER) + 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = 18
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if isinstance(c.value, dt.datetime):
                c.number_format = "DD.MM.YYYY"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(OUT / fname)
    print(f"{fname}: {len(rows)} Personen, {sum(r['Wochenarbeitsstunden'] for r in rows) / 40:g} FTE (inkl. Ein-/Austritte)")

# ---------- Demo-Daten in organigramm.html einsetzen ----------
def cell(v):
    return v.strftime("%Y-%m-%d") if isinstance(v, dt.datetime) else v


# nur die kleine Einheit ist als Beispiel vorausgefüllt
demo = {"klein": [HEADER] + [[cell(r[h]) for h in HEADER] for r in DATASETS["klein"][1]]}
html_path = OUT / "organigramm.html"
html = html_path.read_text(encoding="utf-8")
start, end = "/*DEMO-DATEN-START*/", "/*DEMO-DATEN-ENDE*/"
a, b = html.index(start) + len(start), html.index(end)
html_path.write_text(html[:a] + json.dumps(demo, ensure_ascii=False) + html[b:], encoding="utf-8")
