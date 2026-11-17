# Generator pliku "Kalendarz ładowania akumulatorów.xlsx"
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.comments import Comment
import sys

OUT = sys.argv[1] if len(sys.argv) > 1 else "Kalendarz ladowania akumulatorow.xlsx"
FONT = "Arial"

# --- paleta -----------------------------------------------------------
INK = "1E2328"      # grafit (nagłówki)
INK2 = "2C333A"
ACCENT = "F2C200"   # złoto/żółty akcent
SOFT = "F4F5F7"     # jasne tło kafli
BAND = "F8F9FB"
LINE = "DADDE2"
MUTED = "6B7280"
GREEN = "2E9E6A"; GREEN_L = "E3F4EC"
AMBER = "F0A202"; AMBER_L = "FFF4D6"
RED = "D64545";   RED_L = "FBE5E5"
BLUE_L = "E6F0FB"
INPUT = "FFFBEA"    # pola do wpisywania

def fill(c): return PatternFill("solid", start_color=c, end_color=c)
def font(size=10, bold=False, color=INK, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)
thin = Side(style="thin", color=LINE)
B_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
B_BOT = Border(bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", indent=1)
DATE_FMT = "dd.mm.yyyy"
DNI = '"poniedziałek","wtorek","środa","czwartek","piątek","sobota","niedziela"'

def paint(ws, rng, f=None, fnt=None, al=None, border=None):
    for row in ws[rng]:
        for c in row:
            if f: c.fill = f
            if fnt: c.font = fnt
            if al: c.alignment = al
            if border: c.border = border

wb = Workbook()
P = wb.active; P.title = "Panel"
R = wb.create_sheet("Rejestr ładowań")
K = wb.create_sheet("Kalendarz")
U = wb.create_sheet("Ustawienia")
I = wb.create_sheet("Instrukcja")
for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
P.sheet_properties.tabColor = ACCENT
R.sheet_properties.tabColor = GREEN
K.sheet_properties.tabColor = "3B82F6"
U.sheet_properties.tabColor = MUTED
I.sheet_properties.tabColor = MUTED

LOG_FIRST, LOG_LAST = 6, 1005
LOG_DATE = f"'Rejestr ładowań'!$B${LOG_FIRST}:$B${LOG_LAST}"
LOG_CAR = f"'Rejestr ładowań'!$D${LOG_FIRST}:$D${LOG_LAST}"
CAR_FIRST, NCARS = 13, 12
CAR_LAST = CAR_FIRST + NCARS - 1

# =====================================================================
# USTAWIENIA
# =====================================================================
U.column_dimensions["A"].width = 3
U.column_dimensions["B"].width = 38
U.column_dimensions["C"].width = 16
U.column_dimensions["D"].width = 4
U.column_dimensions["E"].width = 4
U.column_dimensions["F"].width = 14
U.column_dimensions["G"].width = 34
U.column_dimensions["H"].width = 4
U.column_dimensions["I"].width = 16
paint(U, "A1:J2", f=fill(INK))
U["B1"] = "USTAWIENIA"; U["B1"].font = font(16, True, "FFFFFF")
U["B2"] = "Pola na żółtym tle można zmieniać. Reszta liczy się sama."
U["B2"].font = font(9, color="C9CDD3")
U.row_dimensions[1].height = 30

U["B4"] = "Parametr"; U["C4"] = "Wartość"
paint(U, "B4:C4", f=fill(INK2), fnt=font(10, True, "FFFFFF"), al=LEFT)
rows = [
    ("Co ile dni ładować akumulator", 2, "Domyślnie co 2 dni (wymaganie salonu)."),
    ("Czy salon jest otwarty w soboty? (TAK/NIE)", "NIE", "NIE = sobota i niedziela wolne → termin przesuwa się na piątek lub poniedziałek."),
    ("Dni robocze w tygodniu (liczy się samo)", '=IF(C6="TAK","pon–sob","pon–pt")', None),
    ("Osoba odpowiedzialna", "Mateusz", None),
]
for i, (lab, val, note) in enumerate(rows):
    r = 5 + i
    U.cell(r, 2, lab).font = font(10)
    c = U.cell(r, 3, val)
    c.alignment = CENTER
    if i == 2:
        c.font = font(10, color=MUTED); c.fill = fill(SOFT)
    else:
        c.font = font(10, True, "0000FF"); c.fill = fill(INPUT)
    if note: c.comment = Comment(note, "Kalendarz")
    for col in (2, 3): U.cell(r, col).border = B_ALL
    U.row_dimensions[r].height = 22
dv = DataValidation(type="list", formula1='"TAK,NIE"', allow_blank=False)
U.add_data_validation(dv); dv.add("C6")
dvn = DataValidation(type="whole", operator="between", formula1="1", formula2="14")
dvn.error = "Wpisz liczbę dni od 1 do 14."; dvn.errorTitle = "Interwał"
U.add_data_validation(dvn); dvn.add("C5")

U["B11"] = "Zasada wyznaczania terminu"
U["B11"].font = font(11, True)
zasady = [
    "1. Termin = data ostatniego ładowania + liczba dni z ustawień.",
    "2. Jeśli termin wypada w dzień wolny (sobota/niedziela/święto),",
    "    ładowanie przesuwa się na OSTATNI dzień roboczy PRZED nim (np. piątek).",
    "3. Jeśli tym dniem byłby dzień ostatniego ładowania (np. ładowanie w piątek),",
    "    termin przechodzi na PIERWSZY dzień roboczy PO wolnym (poniedziałek).",
]
for i, t in enumerate(zasady):
    U.cell(12 + i, 2, t).font = font(9, color=MUTED)

# święta
U["F4"] = "Dzień wolny"; U["G4"] = "Nazwa"
paint(U, "F4:G4", f=fill(INK2), fnt=font(10, True, "FFFFFF"), al=LEFT)
def easter(y):
    a=y%19;b=y//100;c=y%100;d=b//4;e=b%4;f=(b+8)//25;g=(b-f+1)//3
    h=(19*a+b-d-g+15)%30;i=c//4;k=c%4;l=(32+2*e+2*i-h-k)%7
    m=(a+11*h+22*l)//451;mo=(h+l-7*m+114)//31;da=((h+l-7*m+114)%31)+1
    return dt.date(y,mo,da)
hol = []
for y in (2026, 2027, 2028):
    e = easter(y)
    hol += [(dt.date(y,1,1),"Nowy Rok"),(dt.date(y,1,6),"Trzech Króli"),
            (e+dt.timedelta(1),"Poniedziałek Wielkanocny"),(dt.date(y,5,1),"Święto Pracy"),
            (dt.date(y,5,3),"Święto Konstytucji 3 Maja"),(e+dt.timedelta(60),"Boże Ciało"),
            (dt.date(y,8,15),"Wniebowzięcie NMP"),(dt.date(y,11,1),"Wszystkich Świętych"),
            (dt.date(y,11,11),"Święto Niepodległości"),(dt.date(y,12,24),"Wigilia"),
            (dt.date(y,12,25),"Boże Narodzenie"),(dt.date(y,12,26),"Drugi dzień świąt")]
hol.sort()
HOL_FIRST, HOL_LAST = 5, 5 + 59
for r in range(HOL_FIRST, HOL_LAST + 1):
    for col in (6, 7):
        c = U.cell(r, col); c.fill = fill(INPUT); c.border = B_ALL; c.font = font(9, color="0000FF")
    U.cell(r, 6).number_format = DATE_FMT; U.cell(r, 6).alignment = CENTER
for i, (d, n) in enumerate(hol):
    U.cell(HOL_FIRST + i, 6, d); U.cell(HOL_FIRST + i, 7, n)
U["F3"] = "Święta i dni zamknięcia salonu (dopisuj w wolnych wierszach, np. inwentaryzacja)"
U["F3"].font = font(9, True, MUTED)

# miesiące
U["I4"] = "Miesiące"
paint(U, "I4:I4", f=fill(INK2), fnt=font(10, True, "FFFFFF"), al=LEFT)
MIES = ["Styczeń","Luty","Marzec","Kwiecień","Maj","Czerwiec","Lipiec","Sierpień","Wrzesień","Październik","Listopad","Grudzień"]
for i, m in enumerate(MIES):
    U.cell(5 + i, 9, m).font = font(9, color=MUTED)

D = wb.create_sheet("_dni")
D.sheet_state = "hidden"
D["A1"]="Data"; D["B1"]="Roboczy"; D["C1"]="Nr dnia roboczego"
start = dt.date(2025,1,1); NDAYS = (dt.date(2035,12,31)-start).days+1
for i in range(NDAYS):
    r=i+2
    D.cell(r,1,start+dt.timedelta(i)).number_format=DATE_FMT
    D.cell(r,2,f'=IF(AND(WEEKDAY(A{r},2)<6+(Ustawienia!$C$6="TAK"),COUNTIF(Swieta,A{r})=0),1,0)')
    D.cell(r,3,f'=B{r}' if i==0 else f'=C{r-1}+B{r}')
DL = NDAYS+1
for name, ref in [("DniData", f"_dni!$A$2:$A${DL}"), ("DniRob", f"_dni!$B$2:$B${DL}"),
                  ("DniCum", f"_dni!$C$2:$C${DL}"), ("DniStart", "_dni!$A$2")]:
    wb.defined_names[name] = DefinedName(name, attr_text=ref)
for name, ref in [("Interwal", "Ustawienia!$C$5"),
                  ("Swieta", f"Ustawienia!$F${HOL_FIRST}:$F${HOL_LAST}"),
                  ("Miesiace", "Ustawienia!$I$5:$I$16"),
                  ("Samochody", f"Panel!$F${CAR_FIRST}:$F${CAR_LAST}"),
                  ("Odpowiedzialny", "Ustawienia!$C$8")]:
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# =====================================================================
# PANEL
# =====================================================================
widths = {"A":2.5,"B":6,"C":12,"D":18,"E":20,"F":30,"G":15,"H":15,"I":15,"J":12,"K":17,"L":11,"M":2.5}
for k, v in widths.items(): P.column_dimensions[k].width = v
P.column_dimensions["F"].hidden = True

paint(P, "A1:M3", f=fill(INK))
P.row_dimensions[1].height = 10
P.row_dimensions[2].height = 34
P.row_dimensions[3].height = 22
P["B2"] = "KALENDARZ ŁADOWANIA AKUMULATORÓW"
P["B2"].font = font(20, True, "FFFFFF"); P["B2"].alignment = Alignment(vertical="center")
P["B3"] = "Salon Dacia & Renault  ·  samochody ekspozycyjne"
P["B3"].font = font(10, color=ACCENT); P["B3"].alignment = Alignment(vertical="top")
P.merge_cells("J2:L2"); P["J2"] = "DZIŚ"
P["J2"].font = font(9, True, "9CA3AF"); P["J2"].alignment = Alignment(horizontal="right", vertical="bottom")
P.merge_cells("J3:L3"); P["J3"] = "=TODAY()"
P["J3"].number_format = "dddd, dd.mm.yyyy"
P["J3"].font = font(12, True, "FFFFFF"); P["J3"].alignment = Alignment(horizontal="right", vertical="top")
paint(P, "A4:M4", f=fill(ACCENT)); P.row_dimensions[4].height = 4

# kafle KPI
K_RNG = f"$K${CAR_FIRST}:$K${CAR_LAST}"
tiles = [
    ("B", "D", "SAMOCHODY", f'=COUNTIF($F${CAR_FIRST}:$F${CAR_LAST},"?*")', "w ekspozycji", INK),
    ("E", "G", "DO NAŁADOWANIA DZIŚ", f'=COUNTIF({K_RNG},"*DZIŚ*")', "zaplanowane na dziś", AMBER),
    ("H", "I", "ZALEGŁE", f'=COUNTIF({K_RNG},"*ZALEG*")+COUNTIF({K_RNG},"*NAŁADUJ*")', "wymagają ładowania", RED),
    ("J", "L", "W TERMINIE", f'=COUNTIF({K_RNG},"*OK*")+COUNTIF({K_RNG},"*JUTRO*")', "akumulatory zadbane", GREEN),
]
P.row_dimensions[5].height = 12
P.row_dimensions[6].height = 5
P.row_dimensions[7].height = 20
P.row_dimensions[8].height = 42
P.row_dimensions[9].height = 18
for c1, c2, lab, f_, cap, col in tiles:
    paint(P, f"{c1}6:{c2}6", f=fill(col))
    paint(P, f"{c1}7:{c2}9", f=fill(SOFT))
    for r, val, fnt in [(7, lab, font(8, True, MUTED)), (8, f_, font(28, True, col)), (9, cap, font(8, color=MUTED))]:
        P.merge_cells(f"{c1}{r}:{c2}{r}")
        c = P[f"{c1}{r}"]; c.value = val; c.font = fnt
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    # odstęp między kaflami - prawa biała krawędź
    wside = Side(style="thick", color="FFFFFF")
    for r in range(6, 10):
        P[f"{c2}{r}"].border = Border(right=wside)

# pasek akcji
P.row_dimensions[10].height = 12
P.row_dimensions[11].height = 30
P.merge_cells("B11:E11")
P["B11"] = (f'=HYPERLINK("#\'Rejestr ładowań\'!B"&(COUNT({LOG_DATE})+{LOG_FIRST}),'
            f'"＋  WPISZ ŁADOWANIE")')
paint(P, "B11:E11", f=fill(ACCENT), fnt=font(11, True, INK), al=CENTER)
P.merge_cells("G11:H11"); P["G11"] = "Najbliższe ładowanie:"
P["G11"].font = font(9, True, MUTED); P["G11"].alignment = Alignment(horizontal="right", vertical="center")
P["I11"] = f'=IF(COUNT($H${CAR_FIRST}:$H${CAR_LAST})=0,"",MIN($H${CAR_FIRST}:$H${CAR_LAST}))'
P["I11"].number_format = DATE_FMT; P["I11"].font = font(11, True); P["I11"].alignment = CENTER
P["J11"] = f'=IF(I11="","",CHOOSE(WEEKDAY(I11,2),{DNI}))'
P["J11"].font = font(10, color=MUTED); P["J11"].alignment = LEFT
P.merge_cells("K11:L11"); P["K11"] = '="Odpowiada: "&Odpowiedzialny'
P["K11"].font = font(9, color=MUTED); P["K11"].alignment = Alignment(horizontal="right", vertical="center")
P.row_dimensions[12].height = 8

# nagłówek tabeli (wiersz CAR_FIRST-1 = 12 -> użyjemy 12 jako nagłówek; przesuwamy)
HDR = CAR_FIRST - 1
P.row_dimensions[HDR].height = 32
heads = ["Lp", "Marka", "Model", "Nr rej. / VIN", "Nazwa w rejestrze", "Ostatnie\nładowanie",
         "NASTĘPNE\nładowanie", "Dzień\ntygodnia", "Do terminu", "Status", "Ładowań\nłącznie"]
for i, h in enumerate(heads):
    c = P.cell(HDR, 2 + i, h)
    c.fill = fill(INK); c.font = font(9, True, "FFFFFF"); c.alignment = CENTER
P.cell(HDR, 8).fill = fill(INK2); P.cell(HDR, 8).font = font(9, True, ACCENT)

cars = [("Dacia","Sandero"),("Dacia","Sandero Stepway"),("Dacia","Duster"),("Dacia","Bigster"),
        ("Dacia","Jogger"),("Renault","Clio"),("Renault","Captur"),("Renault","Symbioz"),
        ("Renault","Austral"),("Renault","Rafale")]

for i in range(NCARS):
    r = CAR_FIRST + i
    P.row_dimensions[r].height = 26
    bg = fill("FFFFFF") if i % 2 == 0 else fill(BAND)
    P.cell(r, 2, f'=IF(F{r}="","",COUNTIF($F${CAR_FIRST}:F{r},"?*"))')
    if i < len(cars):
        P.cell(r, 3, cars[i][0]); P.cell(r, 4, cars[i][1])
    P.cell(r, 6, f'=IF(AND(C{r}="",D{r}=""),"",TRIM(C{r}&" "&D{r}&IF(E{r}="",""," · "&E{r})))')
    P.cell(r, 7, f'=IF(F{r}="","",IF(COUNTIF({LOG_CAR},F{r})=0,"",_xlfn.MAXIFS({LOG_DATE},{LOG_CAR},F{r})))')
    due = f"(G{r}+Interwal)"
    cum = lambda x: f"INDEX(DniCum,{x}-DniStart+1)"
    nxt = lambda x: f"INDEX(DniData,MATCH({cum(x)}+1,DniCum,0))"
    prv = lambda x: f"INDEX(DniData,MATCH({cum(x)},DniCum,0))"
    P.cell(r, 8, (f'=IF(F{r}="","",IF(G{r}="",{nxt("(TODAY()-1)")},'
                  f'IF({prv(due)}>G{r},{prv(due)},{nxt(due)})))'))
    P.cell(r, 9, f'=IF(H{r}="","",CHOOSE(WEEKDAY(H{r},2),{DNI}))')
    P.cell(r, 10, f'=IF(H{r}="","",H{r}-TODAY())')
    P.cell(r, 11, (f'=IF(F{r}="","",IF(G{r}="","⚠ NAŁADUJ",IF(H{r}<TODAY(),"✖ ZALEGŁE",'
                   f'IF(H{r}=TODAY(),"⚡ DZIŚ",IF(H{r}=INDEX(DniData,MATCH(INDEX(DniCum,TODAY()-DniStart+1)+1,DniCum,0)),"◔ JUTRO","✔ OK")))))'))
    P.cell(r, 12, f'=IF(F{r}="","",COUNTIF({LOG_CAR},F{r}))')
    for col in range(2, 13):
        c = P.cell(r, col)
        c.fill = bg; c.border = B_BOT; c.alignment = CENTER; c.font = font(10)
    for col in (3, 4, 5):
        P.cell(r, col).alignment = LEFT
        P.cell(r, col).font = font(10, color="0000FF")
    P.cell(r, 4).font = font(10, True, "0000FF")
    P.cell(r, 2).font = font(9, color=MUTED)
    P.cell(r, 7).number_format = DATE_FMT
    P.cell(r, 8).number_format = DATE_FMT; P.cell(r, 8).font = font(11, True)
    P.cell(r, 9).font = font(9, color=MUTED)
    P.cell(r, 10).number_format = '"za "0" d.";"spóźn. "0" d.";"dziś"'
    P.cell(r, 10).font = font(9, color=MUTED)
    P.cell(r, 11).font = font(9, True)
    P.cell(r, 12).font = font(10, color=MUTED)
    for col in (3, 4, 5):
        P.cell(r, col).protection = P.cell(r, col).protection.copy(locked=False)

dvm = DataValidation(type="list", formula1='"Dacia,Renault"', allow_blank=True)
P.add_data_validation(dvm); dvm.add(f"C{CAR_FIRST}:C{CAR_LAST}")

# formatowanie warunkowe statusu
S = f"K{CAR_FIRST}:K{CAR_LAST}"
k1 = f"$K{CAR_FIRST}"
for txt, bgc, fc in [("ZALEG", RED, "FFFFFF"), ("NAŁADUJ", RED, "FFFFFF"), ("DZIŚ", AMBER, INK),
                     ("JUTRO", AMBER_L, "8A6100"), ("OK", GREEN_L, GREEN)]:
    P.conditional_formatting.add(S, FormulaRule(formula=[f'ISNUMBER(SEARCH("{txt}",{k1}))'],
        fill=fill(bgc), font=Font(name=FONT, bold=True, color=fc), stopIfTrue=True))
# podświetlenie terminu
H_ = f"H{CAR_FIRST}:H{CAR_LAST}"
P.conditional_formatting.add(H_, FormulaRule(formula=[f'AND($H{CAR_FIRST}<>"",$H{CAR_FIRST}<TODAY())'],
    font=Font(name=FONT, bold=True, color=RED), fill=fill(RED_L)))
P.conditional_formatting.add(H_, FormulaRule(formula=[f'AND($H{CAR_FIRST}<>"",$H{CAR_FIRST}=TODAY())'],
    font=Font(name=FONT, bold=True, color="8A6100"), fill=fill(AMBER_L)))

# stopka / legenda
L0 = CAR_LAST + 2
P[f"B{L0}"] = "Jak to działa:"
P[f"B{L0}"].font = font(9, True)
legend = [
    "1.  Po naładowaniu auta kliknij żółty przycisk „WPISZ ŁADOWANIE” i dopisz datę oraz samochód z listy.",
    "2.  System sam wyznacza następny termin (co 2 dni), omija soboty, niedziele i święta – przesuwa na piątek lub poniedziałek.",
    "3.  Kolory: czerwony = zaległe / nigdy nie ładowane · pomarańczowy = dziś · jasnożółty = następny dzień roboczy · zielony = w terminie.",
    "4.  Niebieski tekst = pola do edycji (marka, model, nr rej.). Pozostałe pola liczą się automatycznie.",
]
for i, t in enumerate(legend):
    P[f"B{L0+1+i}"] = t; P[f"B{L0+1+i}"].font = font(9, color=MUTED)

P.freeze_panes = f"A{CAR_FIRST}"
P.protection.sheet = True
P.protection.formatCells = False
P.protection.formatColumns = False
P.protection.formatRows = False
P.page_setup.orientation = "landscape"; P.page_setup.fitToWidth = 1; P.page_setup.fitToHeight = 0
P.sheet_properties.pageSetUpPr.fitToPage = True
P.print_area = f"A1:M{L0+5}"

# =====================================================================
# REJESTR ŁADOWAŃ
# =====================================================================
for k, v in {"A":2.5,"B":15,"C":15,"D":34,"E":16,"F":14,"G":40,"H":2.5}.items():
    R.column_dimensions[k].width = v
paint(R, "A1:H2", f=fill(INK))
R.row_dimensions[1].height = 32
R["B1"] = "REJESTR ŁADOWAŃ"; R["B1"].font = font(18, True, "FFFFFF")
R["B1"].alignment = Alignment(vertical="center")
R["B2"] = "Wpisz datę (skrót Ctrl+; wstawia dzisiejszą) i wybierz samochód z listy. Jeden wiersz = jedno ładowanie."
R["B2"].font = font(9, color=ACCENT)
paint(R, "A3:H3", f=fill(ACCENT)); R.row_dimensions[3].height = 4
R.row_dimensions[4].height = 10
heads = ["Data ładowania", "Dzień tygodnia", "Samochód", "Kto ładował", "Napięcie [V]\n(opcjonalnie)", "Uwagi"]
R.row_dimensions[5].height = 32
for i, h in enumerate(heads):
    c = R.cell(5, 2 + i, h); c.fill = fill(INK); c.font = font(9, True, "FFFFFF"); c.alignment = CENTER
for r in range(LOG_FIRST, LOG_LAST + 1):
    R.row_dimensions[r].height = 20
    R.cell(r, 3, f'=IF(B{r}="","",CHOOSE(WEEKDAY(B{r},2),{DNI}))')
    for col in range(2, 8):
        c = R.cell(r, col); c.border = B_BOT; c.font = font(10, color="0000FF"); c.alignment = CENTER
    R.cell(r, 3).font = font(9, color=MUTED)
    R.cell(r, 2).number_format = DATE_FMT
    R.cell(r, 4).alignment = LEFT; R.cell(r, 7).alignment = LEFT
    R.cell(r, 6).number_format = "0.0"
LR = f"B{LOG_FIRST}:G{LOG_LAST}"
R.conditional_formatting.add(LR, FormulaRule(formula=[f'AND($B{LOG_FIRST}<>"",$D{LOG_FIRST}="")'], fill=fill(RED_L)))
R.conditional_formatting.add(LR, FormulaRule(formula=[f'AND($B{LOG_FIRST}="",$D{LOG_FIRST}<>"")'], fill=fill(RED_L)))
R.conditional_formatting.add(LR, FormulaRule(formula=[f'MOD(ROW(),2)=0'], fill=fill(BAND)))
R.conditional_formatting.add(f"F{LOG_FIRST}:F{LOG_LAST}",
    FormulaRule(formula=[f'AND($F{LOG_FIRST}<>"",$F{LOG_FIRST}<12.2)'], font=Font(name=FONT, bold=True, color=RED)))

dvd = DataValidation(type="date", operator="between", formula1="DATE(2024,1,1)", formula2="TODAY()")
dvd.errorTitle = "Nieprawidłowa data"; dvd.error = "Wpisz datę ładowania w formacie dd.mm.rrrr (nie może być z przyszłości)."
dvd.promptTitle = "Data ładowania"; dvd.prompt = "Ctrl + ; wstawia dzisiejszą datę"
dvd.showInputMessage = True; dvd.showErrorMessage = True
R.add_data_validation(dvd); dvd.add(f"B{LOG_FIRST}:B{LOG_LAST}")
dvc = DataValidation(type="list", formula1="=Samochody", allow_blank=True)
dvc.errorTitle = "Wybierz z listy"; dvc.error = "Wybierz samochód z listy (lista pochodzi z zakładki Panel)."
dvc.showErrorMessage = True
R.add_data_validation(dvc); dvc.add(f"D{LOG_FIRST}:D{LOG_LAST}")
dvv = DataValidation(type="decimal", operator="between", formula1="0", formula2="20", allow_blank=True)
dvv.error = "Napięcie podaj w woltach, np. 12,4"; dvv.showErrorMessage = True
R.add_data_validation(dvv); dvv.add(f"F{LOG_FIRST}:F{LOG_LAST}")
R.cell(5, 6).comment = Comment("Opcjonalnie: napięcie zmierzone przed ładowaniem. Poniżej 12,2 V podświetli się na czerwono.", "Kalendarz")
R.freeze_panes = f"A{LOG_FIRST}"
R.auto_filter.ref = f"B5:G{LOG_LAST}"
R.print_title_rows = "5:5"; R.print_area = f"A1:H{LOG_FIRST+59}"
R.page_setup.fitToWidth = 1; R.page_setup.fitToHeight = 0; R.sheet_properties.pageSetUpPr.fitToPage = True

# =====================================================================
# KALENDARZ MIESIĘCZNY
# =====================================================================
K.column_dimensions["A"].width = 2.5
K.column_dimensions["B"].width = 30
from openpyxl.utils import get_column_letter as L
for d in range(31):
    K.column_dimensions[L(3 + d)].width = 4.3
SUMC = L(34)
K.column_dimensions[SUMC].width = 9
K.column_dimensions[L(35)].width = 2.5
LASTC = L(35)
paint(K, f"A1:{LASTC}2", f=fill(INK))
K.row_dimensions[1].height = 32
K["B1"] = "KALENDARZ MIESIĘCZNY"; K["B1"].font = font(18, True, "FFFFFF"); K["B1"].alignment = Alignment(vertical="center")
K["B2"] = "Wybierz miesiąc i rok – kalendarz pokaże wykonane ładowania i najbliższy termin każdego auta."
K["B2"].font = font(9, color=ACCENT)
paint(K, f"A3:{LASTC}3", f=fill(ACCENT)); K.row_dimensions[3].height = 4
K.row_dimensions[5].height = 26
K["B5"] = "Miesiąc  ▸"; K["B5"].font = font(10, True, MUTED); K["B5"].alignment = Alignment(horizontal="right", vertical="center")
K.merge_cells("C5:H5"); K["C5"] = "=INDEX(Miesiace,MONTH(TODAY()))"
K.merge_cells("I5:J5"); K["I5"] = "Rok"; K["I5"].font = font(10, True, MUTED); K["I5"].alignment = Alignment(horizontal="right", vertical="center")
K.merge_cells("K5:N5"); K["K5"] = "=YEAR(TODAY())"
for a in ("C5", "K5"):
    paint(K, a + ":" + ("H5" if a == "C5" else "N5"), f=fill(INPUT), border=B_ALL)
    K[a].font = font(12, True, "0000FF"); K[a].alignment = CENTER
    K[a].protection = K[a].protection.copy(locked=False)
dvM = DataValidation(type="list", formula1="=Miesiace"); K.add_data_validation(dvM); dvM.add("C5")
dvY = DataValidation(type="whole", operator="between", formula1="2024", formula2="2035"); K.add_data_validation(dvY); dvY.add("K5")
K["P5"] = "=DATE(K5,MATCH(C5,Miesiace,0),1)"   # pierwszy dzień miesiąca (pomocnicze)
K["P5"].number_format = '"widok: "mmmm yyyy'; K["P5"].font = font(9, color=MUTED)
K.merge_cells("P5:Y5")
K["P5"].alignment = LEFT

# legenda
K.row_dimensions[6].height = 22
leg = [("Z", "✔", GREEN, "FFFFFF", "naładowano"), ("AB", "⚡", AMBER, INK, "termin"),
       ("AE", "⚡", RED, "FFFFFF", "zaległe")]
from openpyxl.utils import column_index_from_string as CI
for col, sym, bgc, fc, txt in leg:
    c = K[f"{col}5"]; c.value = sym; c.fill = fill(bgc); c.font = font(10, True, fc); c.alignment = CENTER
    t = K.cell(5, CI(col) + 1, txt); t.font = font(8, color=MUTED); t.alignment = Alignment(vertical="center")

DROW, NROW = 9, 8   # wiersz z datami (dzień), wiersz z nazwą dnia
K.row_dimensions[7].height = 8
K.row_dimensions[NROW].height = 18
K.row_dimensions[DROW].height = 22
K.cell(NROW, 2, "").fill = fill(INK)
K.cell(DROW, 2, "Samochód").fill = fill(INK)
K.cell(DROW, 2).font = font(9, True, "FFFFFF"); K.cell(DROW, 2).alignment = LEFT
K.cell(NROW, 34).fill = fill(INK); K.cell(DROW, 34, "Ładowań\nw mies.").fill = fill(INK)
K.cell(DROW, 34).font = font(7, True, "FFFFFF"); K.cell(DROW, 34).alignment = CENTER
for d in range(31):
    col = 3 + d; cl = L(col)
    dc = K.cell(DROW, col, f'=IF(MONTH($P$5+{d})=MONTH($P$5),$P$5+{d},"")')
    dc.number_format = "d"; dc.font = font(9, True, "FFFFFF"); dc.fill = fill(INK); dc.alignment = CENTER
    nc = K.cell(NROW, col, f'=IF({cl}{DROW}="","",CHOOSE(WEEKDAY({cl}{DROW},2),"pn","wt","śr","cz","pt","so","nd"))')
    nc.font = font(7, color="9CA3AF"); nc.fill = fill(INK); nc.alignment = CENTER

KF = DROW + 1
for i in range(NCARS):
    r = KF + i; pr = CAR_FIRST + i
    K.row_dimensions[r].height = 24
    K.cell(r, 2, f'=IF(Panel!F{pr}="","",Panel!F{pr})').font = font(9, True)
    K.cell(r, 2).alignment = LEFT; K.cell(r, 2).border = Border(bottom=thin, right=thin)
    for d in range(31):
        col = 3 + d; cl = L(col)
        f_ = (f'=IF(OR({cl}${DROW}="",$B{r}=""),"",IF(COUNTIFS({LOG_CAR},$B{r},{LOG_DATE},{cl}${DROW})>0,"✔",'
              f'IF({cl}${DROW}=Panel!$H${pr},"⚡","")))')
        c = K.cell(r, col, f_)
        c.alignment = CENTER; c.font = font(10, True); c.border = Border(bottom=thin, right=Side(style="hair", color=LINE))
    sc = K.cell(r, 34, f'=IF($B{r}="","",COUNTIF(C{r}:AG{r},"✔"))')
    sc.alignment = CENTER; sc.font = font(10, True, MUTED); sc.border = Border(bottom=thin, left=thin)

GR = f"C{KF}:AG{KF+NCARS-1}"
top = f"C${DROW}"
K.conditional_formatting.add(GR, FormulaRule(formula=[f'C{KF}="✔"'], fill=fill(GREEN), font=Font(name=FONT, bold=True, color="FFFFFF"), stopIfTrue=True))
K.conditional_formatting.add(GR, FormulaRule(formula=[f'AND(C{KF}="⚡",C${DROW}<TODAY())'], fill=fill(RED), font=Font(name=FONT, bold=True, color="FFFFFF"), stopIfTrue=True))
K.conditional_formatting.add(GR, FormulaRule(formula=[f'C{KF}="⚡"'], fill=fill(AMBER), font=Font(name=FONT, bold=True, color=INK), stopIfTrue=True))
K.conditional_formatting.add(GR, FormulaRule(formula=[f'C${DROW}=TODAY()'], fill=fill(BLUE_L)))
K.conditional_formatting.add(GR, FormulaRule(formula=[f'AND(C${DROW}<>"",INDEX(DniRob,C${DROW}-DniStart+1)=0)'], fill=fill("ECEDEF")))
HR = f"C{DROW}:AG{DROW}"
K.conditional_formatting.add(HR, FormulaRule(formula=[f'C{DROW}=TODAY()'], fill=fill(ACCENT), font=Font(name=FONT, bold=True, color=INK), stopIfTrue=True))
K.conditional_formatting.add(f"C{NROW}:AG{NROW}", FormulaRule(formula=[f'AND(C{DROW}<>"",INDEX(DniRob,C{DROW}-DniStart+1)=0)'], font=Font(name=FONT, color="6B7280", italic=True)))
K.conditional_formatting.add(HR, FormulaRule(formula=[f'AND(C{DROW}<>"",INDEX(DniRob,C{DROW}-DniStart+1)=0)'], fill=fill("3A424B")))

K.freeze_panes = f"C{KF}"
K.protection.sheet = True
K.protection.formatCells = False; K.protection.formatColumns = False; K.protection.formatRows = False
K.page_setup.orientation = "landscape"; K.sheet_properties.pageSetUpPr.fitToPage = True
K.page_setup.fitToWidth = 1; K.page_setup.fitToHeight = 0
K.print_area = f"A1:{LASTC}{KF+NCARS}"
U.print_area = "A1:J64"; U.sheet_properties.pageSetUpPr.fitToPage = True; U.page_setup.fitToWidth = 1; U.page_setup.fitToHeight = 0

# =====================================================================
# INSTRUKCJA
# =====================================================================
I.column_dimensions["A"].width = 2.5; I.column_dimensions["B"].width = 6; I.column_dimensions["C"].width = 100
paint(I, "A1:D2", f=fill(INK)); I.row_dimensions[1].height = 32
I["B1"] = "INSTRUKCJA"; I["B1"].font = font(18, True, "FFFFFF"); I["B1"].alignment = Alignment(vertical="center")
I["B2"] = "Codzienna obsługa zajmuje ok. 1 minuty."; I["B2"].font = font(9, color=ACCENT)
paint(I, "A3:D3", f=fill(ACCENT)); I.row_dimensions[3].height = 4
steps = [
    ("RAZ NA POCZĄTKU", None),
    ("1", "Zakładka Panel: wpisz (niebieski tekst) markę, model i numer rejestracyjny / VIN dla każdego auta z ekspozycji. Jest 12 miejsc."),
    ("2", "Zakładka Ustawienia: sprawdź interwał (2 dni), czy salon pracuje w soboty, oraz listę świąt i dni zamknięcia."),
    ("CODZIENNIE", None),
    ("3", "Otwórz Panel. Auta z czerwonym lub pomarańczowym statusem trzeba dziś naładować."),
    ("4", "Po naładowaniu kliknij żółty przycisk „WPISZ ŁADOWANIE” – przeniesie Cię do pierwszego wolnego wiersza rejestru."),
    ("5", "Wpisz datę (Ctrl + ; wstawia dzisiejszą), wybierz samochód z listy i – jeśli chcesz – napięcie oraz uwagi."),
    ("6", "Wróć do Panelu: status zmieni się na zielony, a kolumna NASTĘPNE pokaże kolejny termin."),
    ("DOBRZE WIEDZIEĆ", None),
    ("•", "Termin = ostatnie ładowanie + 2 dni. Gdy wypada w sobotę → piątek; gdy w niedzielę (po ładowaniu w piątek) → poniedziałek. Święta są omijane tak samo."),
    ("•", "Zmiana nazwy auta w Panelu (np. nowy nr rej.) odłącza starą historię – przy wymianie auta na ekspozycji wpisz je w nowym wierszu lub nadpisz stary."),
    ("•", "Wpis podświetlony na czerwono w rejestrze = brakuje daty albo samochodu."),
    ("•", "Arkusze Panel i Kalendarz są zabezpieczone przed przypadkowym skasowaniem wzorów (bez hasła: Recenzja → Nie chroń arkusza)."),
]
r = 5
for a, b in steps:
    if b is None:
        r += 1
        I.cell(r, 2, a).font = font(10, True, INK)
        paint(I, f"B{r}:C{r}", border=B_BOT)
    else:
        I.cell(r, 2, a).font = font(11, True, AMBER); I.cell(r, 2).alignment = Alignment(horizontal="center", vertical="top")
        I.cell(r, 3, b).font = font(10); I.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="top")
        I.row_dimensions[r].height = 30
    r += 1

wb.active = 0
wb.save(OUT)
print("zapisano", OUT)
