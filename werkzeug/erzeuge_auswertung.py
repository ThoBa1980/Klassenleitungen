"""Erzeugt die Excel-Auswertung für die Klassen-Ampel.

Aufruf:  python erzeuge_auswertung.py            -> leere Vorlage
         python erzeuge_auswertung.py --demo     -> Vorlage mit Beispieldaten
"""
import datetime as dt
import random
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

DEMO = "--demo" in sys.argv
N = 1000          # max. Antworten, die ausgewertet werden
LAST = N + 1      # letzte Zeile in Rohdaten/Daten
FONT = "Arial"

F = lambda **kw: Font(name=FONT, **kw)
FILL_IN = PatternFill("solid", fgColor="FFF2CC")
FILL_HEAD = PatternFill("solid", fgColor="1F4E78")
FILL_GREEN = PatternFill("solid", fgColor="C6EFCE")
FILL_YELLOW = PatternFill("solid", fgColor="FFEB9C")
FILL_RED = PatternFill("solid", fgColor="FFC7CE")
FILL_GREY = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
DATE = "DD.MM.YYYY"
DAYS = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag"]
SLOTS = ["1./2. Stunde", "3./4. Stunde", "5./6. Stunde"]

wb = Workbook()


def head(ws, row, values, col=1):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=col + i, value=v)
        c.font = F(bold=True, color="FFFFFF")
        c.fill = FILL_HEAD
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = F(bold=True, size=16, color="1F4E78")
    if sub:
        ws["A2"] = sub
        ws["A2"].font = F(italic=True, color="595959")


def tidy_chart(ch, xt=None, yt=None):
    ch.roundedCorners = False
    ch.legend = None
    if ch.title is not None:
        pass
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    if xt:
        ch.x_axis.title = xt
    if yt:
        ch.y_axis.title = yt
    ch.height, ch.width = 7.5, 16


# ---------------------------------------------------------------- Übersicht
ov = wb.active
ov.title = "Übersicht"

# ---------------------------------------------------------------- Einstellungen
es = wb.create_sheet("Einstellungen")
title(es, "Einstellungen", "Nur die gelben Felder ändern.")
rows = [
    (4, "Klasse", "7b", None, "Name der Klasse – erscheint in der Übersicht."),
    (5, "Startdatum", dt.date(2026, 11, 30), DATE, "Erster Tag, der gezählt wird."),
    (6, "Enddatum (Halbjahresende)", dt.date(2027, 1, 29), DATE,
     "Letzter Tag, der gezählt wird. PLATZHALTER – bitte an euer Halbjahresende anpassen."),
    (8, "Punkte für GRÜN", 1, "+0;-0;0", "Vorgabe der Klassenleitung."),
    (9, "Punkte für GELB", 0, "+0;-0;0", "Vorgabe der Klassenleitung."),
    (10, "Punkte für ROT", -1, "+0;-0;0", "Vorgabe der Klassenleitung."),
]
for r, label, val, fmt, note in rows:
    es.cell(row=r, column=1, value=label).font = F(bold=True)
    c = es.cell(row=r, column=2, value=val)
    c.font = F(color="0000FF")
    c.fill = FILL_IN
    c.border = BOX
    if fmt:
        c.number_format = fmt
    es.cell(row=r, column=3, value=note).font = F(color="595959")

es["A13"] = "Technik (nicht ändern) – in welcher Spalte von „Rohdaten“ steht was?"
es["A13"].font = F(bold=True, color="808080")
H = "Rohdaten!$A$1:$Z$1"
tech = [
    (14, "Spalte Uhrzeit (Fertigstellung)",
     f'=IFERROR(MATCH("Fertigstellungszeit",{H},0),IFERROR(MATCH("Completion time",{H},0),'
     f'IFERROR(MATCH("Abschlusszeit",{H},0),IFERROR(MATCH("Startzeit",{H},0),3))))'),
    (15, "Spalte Doppelstunde", f'=IFERROR(MATCH("Doppelstunde*",{H},0),0)'),
    (16, "Spalte Ampel", f'=IFERROR(MATCH("Ampel*",{H},0),0)'),
    (17, "Spalte anderes Datum", f'=IFERROR(MATCH("Anderes Datum*",{H},0),0)'),
]
for r, label, fml in tech:
    es.cell(row=r, column=1, value=label).font = F(color="808080")
    es.cell(row=r, column=2, value=fml).font = F(color="808080")
es["A19"] = "Prüfung"
es["A19"].font = F(bold=True)
es["B19"] = ('=IF(COUNTA(Rohdaten!$A$2:$A$20)=0,"Noch keine Antworten eingefügt (Blatt „Rohdaten“).",IF(AND(B15>0,B16>0),"OK – Spalten gefunden",'
             '"FEHLER: Spalten „Doppelstunde“ / „Ampel“ in Rohdaten nicht gefunden. '
             'Fragetitel im Formular prüfen!"))')
es["B19"].font = F(bold=True)
es.column_dimensions["A"].width = 34
es.column_dimensions["B"].width = 16
es.column_dimensions["C"].width = 80
START, END = "Einstellungen!$B$5", "Einstellungen!$B$6"

# ---------------------------------------------------------------- Rohdaten
rd = wb.create_sheet("Rohdaten")
rd["A1"] = "ID"
rd.sheet_properties.tabColor = "FFC000"
rd["A1"].comment = Comment(
    "Hier die komplette Antworttabelle aus Microsoft Forms einfügen "
    "(inklusive Überschriftenzeile) – genau ab Zelle A1. Alte Daten vorher löschen.", "Ampel")

# ---------------------------------------------------------------- Daten (Hilfsblatt)
dd = wb.create_sheet("Daten")
head(dd, 1, ["Zeitstempel", "Anderes Datum", "Datum", "Doppelstunde (Text)", "Doppelstunde Nr",
             "Ampel (Text)", "Farbe Nr (1=Grün 2=Gelb 3=Rot)", "Punkte", "Wochentag Nr",
             "Zählt? (1/0)", "Montag der Woche", "Monat", "Punkte (gezählt)"])
RAW = f"Rohdaten!$A$1:$Z${LAST}"
for r in range(2, LAST + 1):
    def raw(col):
        return f"INDEX({RAW},{r},Einstellungen!$B${col})"
    dd[f"A{r}"] = f'=IF({raw(14)}="","",IFERROR({raw(14)}*1,""))'
    dd[f"B{r}"] = f'=IF(Einstellungen!$B$17=0,"",IF({raw(17)}="","",IFERROR({raw(17)}*1,"")))'
    dd[f"C{r}"] = f'=IF(ISNUMBER(B{r}),INT(B{r}),IF(ISNUMBER(A{r}),INT(A{r}),""))'
    dd[f"D{r}"] = f'=IF(Einstellungen!$B$15=0,"",{raw(15)}&"")'
    dd[f"E{r}"] = (f'=IF(ISNUMBER(SEARCH("1./2.",D{r})),1,IF(ISNUMBER(SEARCH("3./4.",D{r})),2,'
                   f'IF(ISNUMBER(SEARCH("5./6.",D{r})),3,0)))')
    dd[f"F{r}"] = f'=IF(Einstellungen!$B$16=0,"",{raw(16)}&"")'
    dd[f"G{r}"] = (f'=IF(ISNUMBER(SEARCH("Grün",F{r})),1,IF(ISNUMBER(SEARCH("Gelb",F{r})),2,'
                   f'IF(ISNUMBER(SEARCH("Rot",F{r})),3,0)))')
    dd[f"H{r}"] = f'=IF(G{r}=0,"",INDEX(Einstellungen!$B$8:$B$10,G{r}))'
    dd[f"I{r}"] = f'=IF(ISNUMBER(C{r}),WEEKDAY(C{r},2),"")'
    # Zählt nur: im Zeitraum, Mo-Fr, gültige Angaben, und nur der LETZTE Eintrag je Tag+Doppelstunde
    dd[f"J{r}"] = (f'=IF(AND(ISNUMBER(C{r}),E{r}>0,G{r}>0),IF(AND(C{r}>={START},C{r}<={END},I{r}<=5,'
                   f'COUNTIFS($C$2:$C${LAST},C{r},$E$2:$E${LAST},E{r},$A$2:$A${LAST},">"&A{r})=0),1,0),0)')
    dd[f"K{r}"] = f'=IF(J{r}=1,C{r}-I{r}+1,"")'
    dd[f"L{r}"] = f'=IF(J{r}=1,DATE(YEAR(C{r}),MONTH(C{r}),1),"")'
    dd[f"M{r}"] = f'=IF(J{r}=1,H{r},0)'
    for col in "ABCKL":
        dd[f"{col}{r}"].number_format = DATE
    dd[f"A{r}"].number_format = "DD.MM.YYYY HH:MM"
for i, w in enumerate([17, 13, 12, 18, 10, 22, 12, 9, 10, 9, 13, 11, 10], start=1):
    dd.column_dimensions[dd.cell(row=1, column=i).column_letter].width = w
dd.freeze_panes = "A2"
dd.sheet_state = "hidden"

D = lambda col: f"Daten!${col}$2:${col}${LAST}"

# ---------------------------------------------------------------- Wochen
wk = wb.create_sheet("Wochen")
title(wk, "Wochenabrechnung", "Jede Zeile = eine Schulwoche (Mo–Fr). Max. 15 Doppelstunden pro Woche.")
head(wk, 4, ["Nr", "KW", "Montag", "Freitag", "Bewertungen", "Grün", "Gelb", "Rot",
             "Punkte", "Ø pro Doppelstunde", "Abdeckung (von 15)", "Punkte kumuliert"])
WROWS = range(5, 5 + 16)
for i, r in enumerate(WROWS):
    wk[f"A{r}"] = i + 1
    first = f"{START}-WEEKDAY({START},2)+1"
    wk[f"C{r}"] = f'=IF({first}+7*(A{r}-1)>{END},"",{first}+7*(A{r}-1))'
    wk[f"D{r}"] = f'=IF(C{r}="","",C{r}+4)'
    wk[f"B{r}"] = f'=IF(C{r}="","",WEEKNUM(C{r},21))'
    wk[f"E{r}"] = f'=IF(C{r}="","",COUNTIFS({D("K")},C{r},{D("J")},1))'
    for col, n in (("F", 1), ("G", 2), ("H", 3)):
        wk[f"{col}{r}"] = f'=IF(C{r}="","",COUNTIFS({D("K")},C{r},{D("J")},1,{D("G")},{n}))'
    wk[f"I{r}"] = f'=IF(C{r}="","",SUMIFS({D("M")},{D("K")},C{r}))'
    wk[f"J{r}"] = f'=IF(C{r}="","",IF(E{r}>0,I{r}/E{r},"–"))'
    wk[f"K{r}"] = f'=IF(C{r}="","",E{r}/15)'
    wk[f"L{r}"] = f'=IF(C{r}="","",SUM($I$5:I{r}))'
    for col in "CD":
        wk[f"{col}{r}"].number_format = DATE
    wk[f"I{r}"].number_format = "+0;-0;0"
    wk[f"L{r}"].number_format = "+0;-0;0"
    wk[f"J{r}"].number_format = "+0.00;-0.00;0.00"
    wk[f"K{r}"].number_format = "0%"
    for c in range(1, 13):
        wk.cell(row=r, column=c).border = BOX
        wk.cell(row=r, column=c).font = F()
        wk.cell(row=r, column=c).alignment = Alignment(horizontal="center")
wk["A22"] = ("Ferienwochen erscheinen mit 0 Bewertungen und zählen nicht in den Durchschnitt. "
             "Wird eine Doppelstunde mehrfach bewertet, zählt nur der letzte Eintrag (Korrektur möglich).")
wk["A22"].font = F(italic=True, color="595959")
wk.conditional_formatting.add("I5:I20", CellIsRule(operator="greaterThan", formula=["0"], fill=FILL_GREEN))
wk.conditional_formatting.add("I5:I20", CellIsRule(operator="lessThan", formula=["0"], fill=FILL_RED))
for col, w in zip("ABCDEFGHIJKL", [5, 6, 12, 12, 12, 7, 7, 7, 9, 13, 12, 12]):
    wk.column_dimensions[col].width = w
ch = BarChart()
ch.title = "Punkte pro Woche"
ch.add_data(Reference(wk, min_col=9, min_row=5, max_row=20), titles_from_data=False)
ch.set_categories(Reference(wk, min_col=2, min_row=5, max_row=20))
tidy_chart(ch, "Kalenderwoche", "Punkte")
ch.title.overlay = False
wk.add_chart(ch, "N4")
lc = BarChart()
lc.title = "Punktestand (kumuliert)"
lc.add_data(Reference(wk, min_col=12, min_row=5, max_row=20), titles_from_data=False)
lc.set_categories(Reference(wk, min_col=2, min_row=5, max_row=20))
tidy_chart(lc, "Kalenderwoche", "Punkte")
lc.title.overlay = False
wk.add_chart(lc, "N20")
wk.freeze_panes = "A5"

# ---------------------------------------------------------------- Monate
mo = wb.create_sheet("Monate")
title(mo, "Monatsabrechnung")
head(mo, 4, ["Monat", "Bewertungen", "Grün", "Gelb", "Rot", "Punkte", "Ø pro Doppelstunde",
             "Anteil Grün", "Anteil Rot"])
for i, r in enumerate(range(5, 11)):
    first = f"DATE(YEAR({START}),MONTH({START}),1)"
    mo[f"A{r}"] = f'=IF(DATE(YEAR({START}),MONTH({START})+{i},1)>{END},"",DATE(YEAR({START}),MONTH({START})+{i},1))'
    mo[f"B{r}"] = f'=IF(A{r}="","",COUNTIFS({D("L")},A{r},{D("J")},1))'
    for col, n in (("C", 1), ("D", 2), ("E", 3)):
        mo[f"{col}{r}"] = f'=IF(A{r}="","",COUNTIFS({D("L")},A{r},{D("J")},1,{D("G")},{n}))'
    mo[f"F{r}"] = f'=IF(A{r}="","",SUMIFS({D("M")},{D("L")},A{r}))'
    mo[f"G{r}"] = f'=IF(A{r}="","",IF(B{r}>0,F{r}/B{r},"–"))'
    mo[f"H{r}"] = f'=IF(A{r}="","",IF(B{r}>0,C{r}/B{r},"–"))'
    mo[f"I{r}"] = f'=IF(A{r}="","",IF(B{r}>0,E{r}/B{r},"–"))'
    mo[f"A{r}"].number_format = "MMMM YYYY"
    mo[f"F{r}"].number_format = "+0;-0;0"
    mo[f"G{r}"].number_format = "+0.00;-0.00;0.00"
    mo[f"H{r}"].number_format = "0%"
    mo[f"I{r}"].number_format = "0%"
    for c in range(1, 10):
        mo.cell(row=r, column=c).border = BOX
        mo.cell(row=r, column=c).font = F()
        mo.cell(row=r, column=c).alignment = Alignment(horizontal="center")
mo.conditional_formatting.add("F5:F10", CellIsRule(operator="greaterThan", formula=["0"], fill=FILL_GREEN))
mo.conditional_formatting.add("F5:F10", CellIsRule(operator="lessThan", formula=["0"], fill=FILL_RED))
for col, w in zip("ABCDEFGHI", [16, 12, 8, 8, 8, 9, 14, 11, 11]):
    mo.column_dimensions[col].width = w
ch = BarChart()
ch.title = "Ø Punkte pro Doppelstunde je Monat"
ch.add_data(Reference(mo, min_col=7, min_row=5, max_row=10), titles_from_data=False)
ch.set_categories(Reference(mo, min_col=1, min_row=5, max_row=10))
tidy_chart(ch, None, "Ø Punkte")
ch.title.overlay = False
mo.add_chart(ch, "K4")

# ---------------------------------------------------------------- Zeitmuster
zm = wb.create_sheet("Zeitmuster")
title(zm, "Wann läuft es gut – wann schlecht?",
      "Ø Punkte pro Doppelstunde: +1 = immer grün, 0 = gelb/gemischt, −1 = immer rot.")
zm["A4"] = "Ø Punkte je Wochentag und Doppelstunde"
zm["A4"].font = F(bold=True, size=12)
head(zm, 5, ["", *SLOTS, "Ø ganzer Tag"])
for i, day in enumerate(DAYS):
    r = 6 + i
    zm[f"A{r}"] = day
    zm[f"A{r}"].font = F(bold=True)
    for s in range(3):
        col = "BCD"[s]
        zm[f"{col}{r}"] = (f'=IFERROR(SUMIFS({D("M")},{D("I")},{i+1},{D("E")},{s+1},{D("J")},1)/'
                           f'COUNTIFS({D("I")},{i+1},{D("E")},{s+1},{D("J")},1),"–")')
    zm[f"E{r}"] = (f'=IFERROR(SUMIFS({D("M")},{D("I")},{i+1},{D("J")},1)/'
                   f'COUNTIFS({D("I")},{i+1},{D("J")},1),"–")')
zm["A11"] = "Ø ganze Woche"
zm["A11"].font = F(bold=True)
for s in range(3):
    col = "BCD"[s]
    zm[f"{col}11"] = (f'=IFERROR(SUMIFS({D("M")},{D("E")},{s+1},{D("J")},1)/'
                      f'COUNTIFS({D("E")},{s+1},{D("J")},1),"–")')
zm["E11"] = f'=IFERROR(SUM({D("M")})/SUM({D("J")}),"–")'
for r in range(6, 12):
    for c in "BCDE":
        zm[f"{c}{r}"].number_format = "+0.00;-0.00;0.00"
        zm[f"{c}{r}"].alignment = Alignment(horizontal="center")
        zm[f"{c}{r}"].border = BOX
        zm[f"{c}{r}"].font = F(bold=(r == 11 or c == "E"))
zm.conditional_formatting.add("B6:E11", ColorScaleRule(
    start_type="num", start_value=-1, start_color="F8696B",
    mid_type="num", mid_value=0, mid_color="FFEB84",
    end_type="num", end_value=1, end_color="63BE7B"))

zm["A13"] = "Anzahl Bewertungen (je mehr, desto aussagekräftiger)"
zm["A13"].font = F(bold=True, size=12)
head(zm, 14, ["", *SLOTS, "Summe"])
for i, day in enumerate(DAYS):
    r = 15 + i
    zm[f"A{r}"] = day
    zm[f"A{r}"].font = F(bold=True)
    for s in range(3):
        col = "BCD"[s]
        zm[f"{col}{r}"] = f'=COUNTIFS({D("I")},{i+1},{D("E")},{s+1},{D("J")},1)'
    zm[f"E{r}"] = f"=SUM(B{r}:D{r})"
    for c in "BCDE":
        zm[f"{c}{r}"].alignment = Alignment(horizontal="center")
        zm[f"{c}{r}"].border = BOX
        zm[f"{c}{r}"].font = F()

zm["A21"] = "Anteil ROT (wie oft lief es schlecht?)"
zm["A21"].font = F(bold=True, size=12)
head(zm, 22, ["", *SLOTS, "ganzer Tag"])
for i, day in enumerate(DAYS):
    r = 23 + i
    zm[f"A{r}"] = day
    zm[f"A{r}"].font = F(bold=True)
    for s in range(3):
        col = "BCD"[s]
        zm[f"{col}{r}"] = (f'=IFERROR(COUNTIFS({D("I")},{i+1},{D("E")},{s+1},{D("J")},1,{D("G")},3)/'
                           f'COUNTIFS({D("I")},{i+1},{D("E")},{s+1},{D("J")},1),"–")')
    zm[f"E{r}"] = (f'=IFERROR(COUNTIFS({D("I")},{i+1},{D("J")},1,{D("G")},3)/'
                   f'COUNTIFS({D("I")},{i+1},{D("J")},1),"–")')
    for c in "BCDE":
        zm[f"{c}{r}"].number_format = "0%"
        zm[f"{c}{r}"].alignment = Alignment(horizontal="center")
        zm[f"{c}{r}"].border = BOX
        zm[f"{c}{r}"].font = F()
zm.conditional_formatting.add("B23:E27", ColorScaleRule(
    start_type="num", start_value=0, start_color="FFFFFF",
    end_type="num", end_value=0.5, end_color="F8696B"))

# Rangliste (Hilfstabelle für beste/schlechteste Zeit)
zm["G4"] = "Alle 15 Zeitfenster"
zm["G4"].font = F(bold=True, size=12)
head(zm, 5, ["Zeitfenster", "Ø Punkte", "Bewertungen"], col=7)
for i, day in enumerate(DAYS):
    for s in range(3):
        r = 6 + i * 3 + s
        zm[f"G{r}"] = f"{day}, {SLOTS[s]}"
        zm[f"H{r}"] = f"={'BCD'[s]}{6+i}"
        zm[f"I{r}"] = f"={'BCD'[s]}{15+i}"
        zm[f"H{r}"].number_format = "+0.00;-0.00;0.00"
        for c in "GHI":
            zm[f"{c}{r}"].border = BOX
            zm[f"{c}{r}"].font = F()
zm.conditional_formatting.add("H6:H20", ColorScaleRule(
    start_type="num", start_value=-1, start_color="F8696B",
    mid_type="num", mid_value=0, mid_color="FFEB84",
    end_type="num", end_value=1, end_color="63BE7B"))

ch = BarChart()
ch.title = "Ø Punkte je Wochentag"
ch.add_data(Reference(zm, min_col=5, min_row=6, max_row=10), titles_from_data=False)
ch.set_categories(Reference(zm, min_col=1, min_row=6, max_row=10))
tidy_chart(ch, None, "Ø Punkte")
ch.title.overlay = False
ch.height, ch.width = 7, 12
zm.add_chart(ch, "K4")
ch = BarChart()
ch.title = "Ø Punkte je Doppelstunde"
ch.add_data(Reference(zm, min_row=11, min_col=2, max_col=4), from_rows=True, titles_from_data=False)
ch.set_categories(Reference(zm, min_row=5, min_col=2, max_col=4))
tidy_chart(ch, None, "Ø Punkte")
ch.title.overlay = False
ch.height, ch.width = 7, 12
zm.add_chart(ch, "K19")
zm.column_dimensions["A"].width = 16
for c in "BCDE":
    zm.column_dimensions[c].width = 14
zm.column_dimensions["G"].width = 26
zm.column_dimensions["H"].width = 10
zm.column_dimensions["I"].width = 12

# ---------------------------------------------------------------- Übersicht (Schlussabrechnung)
title(ov, "Klassen-Ampel – Übersicht & Schlussabrechnung")
ov["A2"] = ('="Klasse "&Einstellungen!B4&"   |   Zeitraum: "&DAY(Einstellungen!B5)&"."&MONTH(Einstellungen!B5)&"."&YEAR(Einstellungen!B5)'
            '&" – "&DAY(Einstellungen!B6)&"."&MONTH(Einstellungen!B6)&"."&YEAR(Einstellungen!B6)')
ov["A2"].font = F(italic=True, color="595959")
ov["A3"] = '=Einstellungen!B19'
ov["A3"].font = F(bold=True, color="7F7F7F")

kpi = [
    ("A5", "PUNKTESTAND GESAMT", f"=SUM({D('M')})", "+0;-0;0"),
    ("C5", "Bewertete Doppelstunden", f"=SUM({D('J')})", "0"),
    ("E5", "Ø Punkte pro Doppelstunde", f'=IFERROR(A6/C6,"–")', "+0.00;-0.00;0.00"),
]
for cell, lab, fml, fmt in kpi:
    col = cell[0]
    ov[f"{col}5"] = lab
    ov[f"{col}5"].font = F(bold=True, color="595959", size=9)
    ov[f"{col}6"] = fml
    ov[f"{col}6"].font = F(bold=True, size=24, color="1F4E78")
    ov[f"{col}6"].number_format = fmt
    ov[f"{col}6"].alignment = Alignment(horizontal="left")

ov["A8"] = "Verteilung"
ov["A8"].font = F(bold=True, size=12)
head(ov, 9, ["Farbe", "Anzahl", "Anteil"])
for i, (lab, fill) in enumerate((("Grün", FILL_GREEN), ("Gelb", FILL_YELLOW), ("Rot", FILL_RED))):
    r = 10 + i
    ov[f"A{r}"] = lab
    ov[f"A{r}"].fill = fill
    ov[f"B{r}"] = f"=COUNTIFS({D('J')},1,{D('G')},{i+1})"
    ov[f"C{r}"] = f'=IFERROR(B{r}/$C$6,0)'
    ov[f"C{r}"].number_format = "0%"
    for c in "ABC":
        ov[f"{c}{r}"].border = BOX
        ov[f"{c}{r}"].font = F()

ov["A14"] = "Auf einen Blick"
ov["A14"].font = F(bold=True, size=12)
WK = "Wochen!$I$5:$I$20"
facts = [
    ("Beste Woche",
     f'=IF(COUNTIF(Wochen!$E$5:$E$20,">0")=0,"–","KW "&INDEX(Wochen!$B$5:$B$20,MATCH(MAXIFS_PLACEHOLDER,{WK},0))&" ("&IF(MAXIFS_PLACEHOLDER>0,"+","")&MAXIFS_PLACEHOLDER&" Punkte)")'),
    ("Schwächste Woche",
     f'=IF(COUNTIF(Wochen!$E$5:$E$20,">0")=0,"–","KW "&INDEX(Wochen!$B$5:$B$20,MATCH(MINIFS_PLACEHOLDER,{WK},0))&" ("&IF(MINIFS_PLACEHOLDER>0,"+","")&MINIFS_PLACEHOLDER&" Punkte)")'),
    ("Bester Wochentag",
     '=IF(COUNT(Zeitmuster!$E$6:$E$10)=0,"–",INDEX(Zeitmuster!$A$6:$A$10,MATCH(MAX(Zeitmuster!$E$6:$E$10),Zeitmuster!$E$6:$E$10,0))&" (Ø "&IF(MAX(Zeitmuster!$E$6:$E$10)>0,"+","")&FIXED(MAX(Zeitmuster!$E$6:$E$10),2)&")")'),
    ("Schwächster Wochentag",
     '=IF(COUNT(Zeitmuster!$E$6:$E$10)=0,"–",INDEX(Zeitmuster!$A$6:$A$10,MATCH(MIN(Zeitmuster!$E$6:$E$10),Zeitmuster!$E$6:$E$10,0))&" (Ø "&IF(MIN(Zeitmuster!$E$6:$E$10)>0,"+","")&FIXED(MIN(Zeitmuster!$E$6:$E$10),2)&")")'),
    ("Beste Tageszeit",
     '=IF(COUNT(Zeitmuster!$B$11:$D$11)=0,"–",INDEX(Zeitmuster!$B$5:$D$5,MATCH(MAX(Zeitmuster!$B$11:$D$11),Zeitmuster!$B$11:$D$11,0))&" (Ø "&IF(MAX(Zeitmuster!$B$11:$D$11)>0,"+","")&FIXED(MAX(Zeitmuster!$B$11:$D$11),2)&")")'),
    ("Schwächste Tageszeit",
     '=IF(COUNT(Zeitmuster!$B$11:$D$11)=0,"–",INDEX(Zeitmuster!$B$5:$D$5,MATCH(MIN(Zeitmuster!$B$11:$D$11),Zeitmuster!$B$11:$D$11,0))&" (Ø "&IF(MIN(Zeitmuster!$B$11:$D$11)>0,"+","")&FIXED(MIN(Zeitmuster!$B$11:$D$11),2)&")")'),
    ("Bestes Zeitfenster",
     '=IF(COUNT(Zeitmuster!$H$6:$H$20)=0,"–",INDEX(Zeitmuster!$G$6:$G$20,MATCH(MAX(Zeitmuster!$H$6:$H$20),Zeitmuster!$H$6:$H$20,0))&" (Ø "&IF(MAX(Zeitmuster!$H$6:$H$20)>0,"+","")&FIXED(MAX(Zeitmuster!$H$6:$H$20),2)&")")'),
    ("Schwierigstes Zeitfenster",
     '=IF(COUNT(Zeitmuster!$H$6:$H$20)=0,"–",INDEX(Zeitmuster!$G$6:$G$20,MATCH(MIN(Zeitmuster!$H$6:$H$20),Zeitmuster!$H$6:$H$20,0))&" (Ø "&IF(MIN(Zeitmuster!$H$6:$H$20)>0,"+","")&FIXED(MIN(Zeitmuster!$H$6:$H$20),2)&")")'),
]
# Beste/schwächste Woche nur über Wochen MIT Bewertungen -> Hilfsspalte
wk["N1"] = "Hilfe: Punkte nur für Wochen mit Bewertungen"
for r in WROWS:
    wk[f"M{r}"] = f'=IF(C{r}="","",IF(E{r}>0,I{r},""))'
    wk[f"M{r}"].font = F(color="FFFFFF")
wk["N1"] = None
WKH = "Wochen!$M$5:$M$20"
for i, (lab, fml) in enumerate(facts):
    r = 15 + i
    fml = fml.replace("MAXIFS_PLACEHOLDER", f"MAX({WKH})").replace("MINIFS_PLACEHOLDER", f"MIN({WKH})")
    fml = fml.replace(f"MATCH(MAX({WKH}),{WK},0)", f"MATCH(MAX({WKH}),{WKH},0)")
    fml = fml.replace(f"MATCH(MIN({WKH}),{WK},0)", f"MATCH(MIN({WKH}),{WKH},0)")
    ov[f"A{r}"] = lab
    ov[f"A{r}"].font = F(bold=True)
    ov[f"C{r}"] = fml
    ov[f"C{r}"].font = F()
ov["A24"] = ("Hinweis: Aussagen zu einzelnen Zeitfenstern sind erst nach mehreren Wochen belastbar "
             "(Anzahl Bewertungen siehe Blatt „Zeitmuster“).")
ov["A24"].font = F(italic=True, color="595959", size=9)
ov.column_dimensions["A"].width = 26
ov.column_dimensions["B"].width = 10
ov.column_dimensions["C"].width = 26
ov.column_dimensions["D"].width = 4
ov.column_dimensions["E"].width = 26

pc = BarChart()
pc.type = "bar"
pc.title = "Grün / Gelb / Rot"
pc.add_data(Reference(ov, min_col=2, min_row=10, max_row=12), titles_from_data=False)
pc.set_categories(Reference(ov, min_col=1, min_row=10, max_row=12))
tidy_chart(pc)
pc.title.overlay = False
from openpyxl.chart.series import DataPoint
for idx, color in enumerate(["63BE7B", "FFD966", "F8696B"]):
    pt = DataPoint(idx=idx)
    pt.graphicalProperties.solidFill = color
    pc.series[0].dPt.append(pt)
pc.height, pc.width = 6, 12
ov.add_chart(pc, "G4")
lc = BarChart()
lc.title = "Punktestand über die Wochen"
lc.add_data(Reference(wk, min_col=12, min_row=5, max_row=20), titles_from_data=False)
lc.set_categories(Reference(wk, min_col=2, min_row=5, max_row=20))
tidy_chart(lc, "KW", "Punkte")
lc.title.overlay = False
lc.height, lc.width = 6.5, 12
ov.add_chart(lc, "G17")

# ---------------------------------------------------------------- Anleitung
an = wb.create_sheet("Anleitung", 0)
title(an, "So funktioniert die Auswertung")
steps = [
    "1.  Einstellungen: Klasse, Startdatum und Enddatum (Halbjahresende) in die gelben Felder eintragen.",
    "2.  In Microsoft Forms: Antworten → „Ergebnisse in Excel öffnen“.",
    "3.  In dieser Excel-Tabelle alles markieren (Strg+A) und kopieren (Strg+C).",
    "4.  Hier im Blatt „Rohdaten“ alte Inhalte löschen, Zelle A1 anklicken, einfügen (Strg+V).",
    "5.  Fertig – „Übersicht“, „Wochen“, „Monate“ und „Zeitmuster“ rechnen automatisch.",
    "",
    "Regeln:",
    "•  Grün = +1, Gelb = 0, Rot = −1 Punkt (änderbar unter „Einstellungen“).",
    "•  Gezählt werden nur Mo–Fr innerhalb von Start- und Enddatum.",
    "•  Wird dieselbe Doppelstunde am selben Tag mehrfach bewertet, zählt nur der LETZTE Eintrag.",
    "   → Wer sich verdrückt hat, schickt einfach noch einmal die richtige Farbe ab.",
    "•  Wird nachträglich bewertet, das Feld „Anderes Datum“ im Formular ausfüllen – sonst gilt das Absendedatum.",
    "•  Ferienwochen zeigen 0 Bewertungen und verfälschen den Durchschnitt nicht.",
    "",
    "Blätter:",
    "•  Übersicht – Schlussabrechnung: Gesamtpunkte, Verteilung, beste/schwächste Woche, Tag und Tageszeit.",
    "•  Wochen – Wochenabrechnung mit Diagrammen.",
    "•  Monate – Monatsabrechnung.",
    "•  Zeitmuster – Farbtabelle Wochentag × Doppelstunde (grün = läuft gut, rot = läuft schlecht).",
    "•  Einstellungen – gelbe Felder anpassen; unten steht eine automatische Prüfung.",
    "•  Rohdaten – hier kommen die Forms-Antworten hinein.",
]
for i, s in enumerate(steps):
    c = an.cell(row=3 + i, column=1, value=s)
    c.font = F(bold=s.endswith(":"))
an.column_dimensions["A"].width = 110
an.sheet_properties.tabColor = "1F4E78"
ov.sheet_properties.tabColor = "70AD47"
es.sheet_properties.tabColor = "FFC000"

# ---------------------------------------------------------------- Demo-Daten
if DEMO:
    random.seed(7)
    rd.append  # noqa
    headers = ["ID", "Startzeit", "Fertigstellungszeit", "E-Mail", "Name",
               "Doppelstunde", "Ampel", "Anderes Datum (nur ausfüllen, wenn nicht heute)"]
    for i, h in enumerate(headers, start=1):
        rd.cell(row=1, column=i, value=h)
    colors = ["🟢 Grün – lief gut", "🟡 Gelb – okay", "🔴 Rot – lief schlecht"]
    rid = 1
    day = dt.date(2026, 11, 30)
    end = dt.date(2027, 1, 29)
    ferien = (dt.date(2026, 12, 21), dt.date(2027, 1, 1))
    slot_time = {0: (9, 35), 1: (11, 25), 2: (13, 15)}
    while day <= end:
        if day.weekday() < 5 and not (ferien[0] <= day <= ferien[1]):
            for s in range(3):
                if random.random() < 0.12:  # vergessen
                    continue
                score = 0.35 - 0.35 * (s == 2) - 0.25 * (day.weekday() == 4) + 0.2 * (s == 0) \
                    - 0.3 * (day.weekday() == 0 and s == 0) + random.gauss(0, 0.55)
                c = 0 if score > 0.3 else (1 if score > -0.25 else 2)
                h, m = slot_time[s]
                ts = dt.datetime(day.year, day.month, day.day, h, m) + dt.timedelta(minutes=random.randint(0, 20))
                late = random.random() < 0.08
                if late:
                    ts = ts + dt.timedelta(days=1 if day.weekday() < 4 else 3)
                rd.append([rid, ts - dt.timedelta(seconds=25), ts, "anonym", "", SLOTS[s], colors[c],
                           day if late else None])
                rid += 1
        day += dt.timedelta(days=1)
    # Korrektur-Beispiel: erst Rot, dann (Korrektur) Grün
    t = dt.datetime(2026, 12, 2, 9, 40)
    rd.append([rid, t, t, "anonym", "", SLOTS[0], colors[2], None]); rid += 1
    t2 = t + dt.timedelta(minutes=1)
    rd.append([rid, t2, t2, "anonym", "", SLOTS[0], colors[0], None])
    for row in rd.iter_rows(min_row=2):
        row[1].number_format = row[2].number_format = "DD.MM.YYYY HH:MM:SS"
        row[7].number_format = DATE
    rd.column_dimensions["B"].width = rd.column_dimensions["C"].width = 20
    rd.column_dimensions["F"].width = 14
    rd.column_dimensions["G"].width = 22

for ws in wb.worksheets:
    ws.sheet_view.showGridLines = ws.title in ("Rohdaten",)

out = sys.argv[sys.argv.index("--out") + 1] if "--out" in sys.argv else (
    "Klassenampel_Auswertung_DEMO.xlsx" if DEMO else "Klassenampel_Auswertung.xlsx")
wb.save(out)
print("gespeichert:", out)
