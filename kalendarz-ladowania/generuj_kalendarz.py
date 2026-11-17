# Generator: Kalendarz ładowania akumulatorów (salon Dacia / Renault)
import datetime as dt, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L

OUT = sys.argv[1] if len(sys.argv) > 1 else "Kalendarz_ladowania_akumulatorow.xlsx"
FONT = "Arial"
INK, INK2, ACCENT = "1E2328", "2C333A", "F2C200"
SOFT, BAND, LINE, MUTED = "F4F5F7", "F8F9FB", "DADDE2", "6B7280"
GREEN, GREEN_L = "2E9E6A", "E3F4EC"
AMBER, AMBER_L = "F0A202", "FFE7B3"
RED, RED_L = "D64545", "FBE0E0"
BLUE_L, INPUT = "E6F0FB", "FFFBEA"
DATE_FMT = "dd.mm.yyyy"
DNI = '"poniedziałek","wtorek","środa","czwartek","piątek","sobota","niedziela"'
MIES = ["Styczeń","Luty","Marzec","Kwiecień","Maj","Czerwiec","Lipiec","Sierpień","Wrzesień","Październik","Listopad","Grudzień"]

def fill(c): return PatternFill("solid", start_color=c, end_color=c)
def font(size=10, bold=False, color=INK, italic=False): return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)
thin = Side(style="thin", color=LINE)
B_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
B_BOT = Border(bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", indent=1)
UNLOCK = Protection(locked=False)

def paint(ws, rng, f=None, fnt=None, al=None, border=None):
    for row in ws[rng]:
        for c in row:
            if f: c.fill = f
            if fnt: c.font = fnt
            if al: c.alignment = al
            if border: c.border = border

def header_band(ws, last_col, title, subtitle):
    paint(ws, f"A1:{last_col}2", f=fill(INK))
    ws.row_dimensions[1].height = 34; ws.row_dimensions[2].height = 20
    ws["B1"] = title; ws["B1"].font = font(18, True, "FFFFFF"); ws["B1"].alignment = Alignment(vertical="center")
    ws["B2"] = subtitle; ws["B2"].font = font(9, color=ACCENT); ws["B2"].alignment = Alignment(vertical="top")
    paint(ws, f"A3:{last_col}3", f=fill(ACCENT)); ws.row_dimensions[3].height = 4

wb = Workbook()
S = wb.active; S.title = "Samochody"
K = wb.create_sheet("Kalendarz")
U = wb.create_sheet("Ustawienia")
I = wb.create_sheet("Instrukcja")
D = wb.create_sheet("_dni"); D.sheet_state = "hidden"
for ws in (S, K, U, I): ws.sheet_view.showGridLines = False
S.sheet_properties.tabColor = ACCENT; K.sheet_properties.tabColor = GREEN
U.sheet_properties.tabColor = MUTED; I.sheet_properties.tabColor = MUTED

NCARS = 30
S_FIRST = 13; S_LAST = S_FIRST + NCARS - 1           # wiersze aut w "Samochody"
K_HDR = 8; K_CNT = 9; K_FIRST = 10; K_LAST = K_FIRST + NCARS - 1   # wiersze w "Kalendarz"
C0 = 5                                                # pierwsza kolumna dat (E)

# dni robocze pn-pt w siatce kalendarza
d0, d1 = dt.date(2026, 11, 2), dt.date(2028, 12, 29)
DATES = [d0 + dt.timedelta(i) for i in range((d1 - d0).days + 1) if (d0 + dt.timedelta(i)).weekday() < 5]
CL = L(C0 + len(DATES) - 1)                            # ostatnia kolumna dat
GRID = f"Kalendarz!$E${K_FIRST}:${CL}${K_LAST}"
HDR = f"Kalendarz!$E${K_HDR}:${CL}${K_HDR}"

# =====================================================================
# USTAWIENIA
# =====================================================================
for k, v in {"A": 2.5, "B": 40, "C": 16, "D": 4, "E": 14, "F": 32}.items(): U.column_dimensions[k].width = v
header_band(U, "G", "USTAWIENIA", "Pola na żółtym tle można zmieniać. Reszta liczy się sama.")
U["B5"] = "Parametr"; U["C5"] = "Wartość"
paint(U, "B5:C5", f=fill(INK2), fnt=font(10, True, "FFFFFF"), al=LEFT)
for r, (lab, val, note) in enumerate([
        ("Co ile dni ładować akumulator (2 lub 3)", 2, "Wybierz z listy: 2 lub 3 dni."),
        ("Dni ładowania", "poniedziałek – piątek", None),
        ("Osoba odpowiedzialna", "Mateusz", None)], start=6):
    U.cell(r, 2, lab).font = font(10)
    c = U.cell(r, 3, val); c.alignment = CENTER
    if r == 7: c.font = font(9, color=MUTED); c.fill = fill(SOFT)
    else: c.font = font(10, True, "0000FF"); c.fill = fill(INPUT)
    if note: c.comment = Comment(note, "Kalendarz")
    for col in (2, 3): U.cell(r, col).border = B_ALL
    U.row_dimensions[r].height = 22
dv = DataValidation(type="list", formula1='"2,3"', allow_blank=False); U.add_data_validation(dv); dv.add("C6")

U["B11"] = "Jak wyznaczany jest termin"; U["B11"].font = font(11, True)
for i, t in enumerate([
        "1. Termin = ostatnie ładowanie + 2 dni (lub 3, jeśli tak ustawisz).",
        "2. Ładujemy tylko od poniedziałku do piątku. Jeśli termin wypada w weekend",
        "    lub święto – przesuwa się na ostatni dzień roboczy PRZED nim (np. piątek).",
        "3. Jeśli tym dniem byłby dzień samego ładowania (np. ładowanie w piątek),",
        "    termin przechodzi na pierwszy dzień roboczy PO wolnym (poniedziałek)."]):
    U.cell(12 + i, 2, t).font = font(9, color=MUTED)

U["E4"] = "Święta i dni bez ładowania (dopisuj w wolnych wierszach)"; U["E4"].font = font(9, True, MUTED)
U["E5"] = "Data"; U["F5"] = "Nazwa"
paint(U, "E5:F5", f=fill(INK2), fnt=font(10, True, "FFFFFF"), al=LEFT)
def easter(y):
    a=y%19;b=y//100;c=y%100;d=b//4;e=b%4;f=(b+8)//25;g=(b-f+1)//3
    h=(19*a+b-d-g+15)%30;i=c//4;k=c%4;l=(32+2*e+2*i-h-k)%7
    m=(a+11*h+22*l)//451;mo=(h+l-7*m+114)//31;da=((h+l-7*m+114)%31)+1
    return dt.date(y,mo,da)
hol = []
for y in (2026, 2027, 2028):
    e = easter(y)
    hol += [(dt.date(y,1,1),"Nowy Rok"),(dt.date(y,1,6),"Trzech Króli"),(e+dt.timedelta(1),"Poniedziałek Wielkanocny"),
            (dt.date(y,5,1),"Święto Pracy"),(dt.date(y,5,3),"Święto Konstytucji 3 Maja"),(e+dt.timedelta(60),"Boże Ciało"),
            (dt.date(y,8,15),"Wniebowzięcie NMP"),(dt.date(y,11,1),"Wszystkich Świętych"),(dt.date(y,11,11),"Święto Niepodległości"),
            (dt.date(y,12,24),"Wigilia"),(dt.date(y,12,25),"Boże Narodzenie"),(dt.date(y,12,26),"Drugi dzień świąt")]
hol.sort()
H1, H2 = 6, 70
for r in range(H1, H2 + 1):
    for col in (5, 6):
        c = U.cell(r, col); c.fill = fill(INPUT); c.border = B_ALL; c.font = font(9, color="0000FF")
    U.cell(r, 5).number_format = DATE_FMT; U.cell(r, 5).alignment = CENTER
for i, (d, n) in enumerate(hol):
    U.cell(H1 + i, 5, d); U.cell(H1 + i, 6, n)

# ukryty kalendarz dni roboczych (pomocniczy)
D["A1"], D["B1"], D["C1"] = "Data", "Roboczy", "Nr dnia roboczego"
start = dt.date(2025, 1, 1); ND = (dt.date(2035, 12, 31) - start).days + 1
for i in range(ND):
    r = i + 2
    D.cell(r, 1, start + dt.timedelta(i)).number_format = DATE_FMT
    D.cell(r, 2, f'=IF(AND(WEEKDAY(A{r},2)<6,COUNTIF(Swieta,A{r})=0),1,0)')
    D.cell(r, 3, f"=B{r}" if i == 0 else f"=C{r-1}+B{r}")
DL = ND + 1
for name, ref in [("Interwal", "Ustawienia!$C$6"), ("Odpowiedzialny", "Ustawienia!$C$8"),
                  ("Swieta", f"Ustawienia!$E${H1}:$E${H2}"),
                  ("DniData", f"_dni!$A$2:$A${DL}"), ("DniRob", f"_dni!$B$2:$B${DL}"),
                  ("DniCum", f"_dni!$C$2:$C${DL}"), ("DniStart", "_dni!$A$2"),
                  ("KalendarzSiatka", GRID)]:
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

TODAY_LINK = (f'HYPERLINK("#Kalendarz!"&ADDRESS({K_FIRST},IFERROR(MATCH(TODAY(),{HDR},1),1)+{C0-1}),')

# =====================================================================
# SAMOCHODY
# =====================================================================
for k, v in {"A":2.5,"B":5,"C":12,"D":20,"E":24,"F":14,"G":15,"H":14,"I":12,"J":16,"K":10,"L":2.5}.items():
    S.column_dimensions[k].width = v
paint(S, "A1:L3", f=fill(INK))
S.row_dimensions[1].height = 10; S.row_dimensions[2].height = 34; S.row_dimensions[3].height = 22
S["B2"] = "KALENDARZ ŁADOWANIA AKUMULATORÓW"; S["B2"].font = font(20, True, "FFFFFF"); S["B2"].alignment = Alignment(vertical="center")
S["B3"] = "Salon Dacia & Renault  ·  samochody ekspozycyjne"; S["B3"].font = font(10, color=ACCENT); S["B3"].alignment = Alignment(vertical="top")
S.merge_cells("I2:K2"); S["I2"] = "DZIŚ"
S["I2"].font = font(9, True, "9CA3AF"); S["I2"].alignment = Alignment(horizontal="right", vertical="bottom")
S.merge_cells("I3:K3"); S["I3"] = "=TODAY()"; S["I3"].number_format = "dddd, dd.mm.yyyy"
S["I3"].font = font(12, True, "FFFFFF"); S["I3"].alignment = Alignment(horizontal="right", vertical="top")
paint(S, "A4:L4", f=fill(ACCENT)); S.row_dimensions[4].height = 4

ST = f"$J${S_FIRST}:$J${S_LAST}"
tiles = [("B","D","SAMOCHODY", f'=COUNTIF({ST},"?*")', "na liście", INK),
         ("E","F","DO NAŁADOWANIA DZIŚ", f'=COUNTIF({ST},"*DZIŚ*")', "termin przypada dziś", AMBER),
         ("G","H","ZALEGŁE", f'=COUNTIF({ST},"*ZALEG*")+COUNTIF({ST},"*NAŁADUJ*")', "naładuj jak najszybciej", RED),
         ("I","K","W TERMINIE", f'=COUNTIF({ST},"*OK*")+COUNTIF({ST},"*JUTRO*")', "akumulatory zadbane", GREEN)]
for r, h in ((5, 12), (6, 5), (7, 20), (8, 42), (9, 18)): S.row_dimensions[r].height = h
for c1, c2, lab, f_, cap, col in tiles:
    paint(S, f"{c1}6:{c2}6", f=fill(col)); paint(S, f"{c1}7:{c2}9", f=fill(SOFT))
    for r, val, fnt in ((7, lab, font(8, True, MUTED)), (8, f_, font(28, True, col)), (9, cap, font(8, color=MUTED))):
        S.merge_cells(f"{c1}{r}:{c2}{r}"); c = S[f"{c1}{r}"]; c.value = val; c.font = fnt
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for r in range(6, 10): S[f"{c2}{r}"].border = Border(right=Side(style="thick", color="FFFFFF"))

S.row_dimensions[10].height = 12; S.row_dimensions[11].height = 30
S.merge_cells("B11:E11"); S["B11"] = "=" + TODAY_LINK + '"📅  OTWÓRZ KALENDARZ – DZIŚ")'
paint(S, "B11:E11", f=fill(ACCENT), fnt=font(11, True, INK), al=CENTER)
S["F11"] = "Najbliższe:"; S["F11"].font = font(9, True, MUTED); S["F11"].alignment = Alignment(horizontal="right", vertical="center")
S["G11"] = f'=IF(COUNT($G${S_FIRST}:$G${S_LAST})=0,"",MIN($G${S_FIRST}:$G${S_LAST}))'
S["G11"].number_format = DATE_FMT; S["G11"].font = font(11, True); S["G11"].alignment = CENTER
S["H11"] = f'=IF(G11="","",CHOOSE(WEEKDAY(G11,2),{DNI}))'; S["H11"].font = font(10, color=MUTED); S["H11"].alignment = LEFT
S.merge_cells("I11:K11"); S["I11"] = '="Odpowiada: "&Odpowiedzialny&"  ·  co "&Interwal&" dni"'
S["I11"].font = font(9, color=MUTED); S["I11"].alignment = Alignment(horizontal="right", vertical="center")

HR = S_FIRST - 1; S.row_dimensions[HR].height = 34
for i, h in enumerate(["Lp", "Marka", "Model", "Numer VIN", "Ostatnie\nładowanie", "NASTĘPNE\nładowanie",
                       "Dzień\ntygodnia", "Do terminu", "Status", "Ładowań\nłącznie"]):
    c = S.cell(HR, 2 + i, h); c.fill = fill(INK); c.font = font(9, True, "FFFFFF"); c.alignment = CENTER
S.cell(HR, 7).fill = fill(INK2); S.cell(HR, 7).font = font(9, True, ACCENT)
for col in (3, 4, 5):   # kolumny do wpisywania – oznaczenie w nagłówku
    S.cell(HR, col).fill = fill(INK2)

example = [("Dacia", "Sandero", "UU1XXXXXXXX000001")]
cum = lambda x: f"INDEX(DniCum,{x}-DniStart+1)"
nxt = lambda x: f"INDEX(DniData,MATCH({cum(x)}+1,DniCum,0))"
prv = lambda x: f"INDEX(DniData,MATCH({cum(x)},DniCum,0))"
for i in range(NCARS):
    r = S_FIRST + i; kr = K_FIRST + i
    S.row_dimensions[r].height = 24
    bg = fill("FFFFFF") if i % 2 == 0 else fill(BAND)
    row = f"Kalendarz!$E${kr}:${CL}${kr}"
    has = f"COUNTA(C{r}:E{r})=0"
    charged = f"(({row}<>\"\")*({row}<>FALSE)*({HDR}<=TODAY()))"
    S.cell(r, 2, f'=IF({has},"",{i+1})')
    if i < len(example):
        for j, v in enumerate(example[i]): S.cell(r, 3 + j, v)
    S.cell(r, 6, f'=IF({has},"",IF(SUMPRODUCT({charged})=0,"",SUMPRODUCT(MAX({charged}*{HDR}))))')
    due = f"(F{r}+Interwal)"
    S.cell(r, 7, f'=IF({has},"",IF(F{r}="",{nxt("(TODAY()-1)")},IF({prv(due)}>F{r},{prv(due)},{nxt(due)})))')
    S.cell(r, 8, f'=IF(G{r}="","",CHOOSE(WEEKDAY(G{r},2),{DNI}))')
    S.cell(r, 9, f'=IF(G{r}="","",G{r}-TODAY())')
    S.cell(r, 10, (f'=IF({has},"",IF(F{r}="","⚠ NAŁADUJ",IF(G{r}<TODAY(),"✖ ZALEGŁE",IF(G{r}=TODAY(),"⚡ DZIŚ",'
                   f'IF(G{r}={nxt("TODAY()")},"◔ JUTRO","✔ OK")))))'))
    S.cell(r, 11, f'=IF({has},"",SUMPRODUCT({charged}))')
    for col in range(2, 12):
        c = S.cell(r, col); c.fill = bg; c.border = B_BOT; c.alignment = CENTER; c.font = font(10)
    for col in (3, 4, 5):
        c = S.cell(r, col); c.alignment = LEFT; c.font = font(10, col == 4, "0000FF"); c.protection = UNLOCK
    S.cell(r, 5).font = Font(name="Consolas", size=10, color="0000FF")
    S.cell(r, 2).font = font(9, color=MUTED)
    S.cell(r, 6).number_format = DATE_FMT
    S.cell(r, 7).number_format = DATE_FMT; S.cell(r, 7).font = font(11, True)
    S.cell(r, 8).font = font(9, color=MUTED)
    S.cell(r, 9).number_format = '"za "0" d.";"spóźn. "0" d.";"dziś"'; S.cell(r, 9).font = font(9, color=MUTED)
    S.cell(r, 10).font = font(9, True)
    S.cell(r, 11).font = font(10, color=MUTED)

dvm = DataValidation(type="list", formula1='"Dacia,Renault"', allow_blank=True); S.add_data_validation(dvm); dvm.add(f"C{S_FIRST}:C{S_LAST}")
dvvin = DataValidation(type="textLength", operator="equal", formula1="17", allow_blank=True, errorStyle="warning")
dvvin.errorTitle = "Numer VIN"; dvvin.error = "VIN ma zwykle 17 znaków. Na pewno zapisać?"; dvvin.showErrorMessage = True
S.add_data_validation(dvvin); dvvin.add(f"E{S_FIRST}:E{S_LAST}")
k1 = f"$J{S_FIRST}"
for txt, bgc, fc in [("ZALEG", RED, "FFFFFF"), ("NAŁADUJ", RED, "FFFFFF"), ("DZIŚ", AMBER, INK),
                     ("JUTRO", "FFF4D6", "8A6100"), ("OK", GREEN_L, GREEN)]:
    S.conditional_formatting.add(f"J{S_FIRST}:J{S_LAST}", FormulaRule(formula=[f'ISNUMBER(SEARCH("{txt}",{k1}))'],
        fill=fill(bgc), font=Font(name=FONT, bold=True, color=fc), stopIfTrue=True))
S.conditional_formatting.add(f"G{S_FIRST}:G{S_LAST}", FormulaRule(formula=[f'AND($G{S_FIRST}<>"",$G{S_FIRST}<TODAY())'],
    font=Font(name=FONT, bold=True, color=RED), fill=fill(RED_L)))
S.conditional_formatting.add(f"G{S_FIRST}:G{S_LAST}", FormulaRule(formula=[f'AND($G{S_FIRST}<>"",$G{S_FIRST}=TODAY())'],
    font=Font(name=FONT, bold=True, color="8A6100"), fill=fill("FFF4D6")))

L0 = S_LAST + 2
S[f"B{L0}"] = "Jak to działa:"; S[f"B{L0}"].font = font(9, True)
for i, t in enumerate([
        "1.  Tutaj wpisujesz tylko auta: markę, model i VIN (niebieski tekst). Kolejne auto = kolejny wiersz.",
        "2.  Ładowania zaznaczasz w zakładce Kalendarz: w kratce danego dnia przy danym aucie wpisz x i Enter (pokaże się ✔).",
        "3.  Data ostatniego i następnego ładowania liczą się same: +2 dni, tylko pn–pt, z pominięciem świąt.",
        "4.  Sprzedane auto: usuń markę, model i VIN – ALE też jego ✔ w kalendarzu (albo wpisz nowe auto w wolnym wierszu)."]):
    S[f"B{L0+1+i}"] = t; S[f"B{L0+1+i}"].font = font(9, color=MUTED)
S.freeze_panes = f"A{S_FIRST}"
S.protection.sheet = True; S.protection.formatColumns = False; S.protection.formatRows = False; S.protection.formatCells = False
S.page_setup.orientation = "portrait"; S.sheet_properties.pageSetUpPr.fitToPage = True
S.page_setup.fitToWidth = 1; S.page_setup.fitToHeight = 0; S.print_area = f"A1:L{L0+5}"

# =====================================================================
# KALENDARZ (siatka do klikania)
# =====================================================================
K.column_dimensions["A"].width = 2.5; K.column_dimensions["B"].width = 26
K.column_dimensions["C"].width = 10; K.column_dimensions["D"].width = 11
for j in range(len(DATES)): K.column_dimensions[L(C0 + j)].width = 4.6
header_band(K, CL, "KALENDARZ ŁADOWAŃ",
            "Pomarańczowa kratka = naładuj dziś. Po naładowaniu wpisz w kratce x i Enter – pojawi się ✔, a termin przesunie się sam.")
K.row_dimensions[4].height = 8; K.row_dimensions[5].height = 26
K.merge_cells("B5:D5"); K["B5"] = "=" + TODAY_LINK + '"⟶  PRZEJDŹ DO DZIŚ")'
paint(K, "B5:D5", f=fill(ACCENT), fnt=font(10, True, INK), al=CENTER)
leg = [(C0+1, "✔", GREEN, "FFFFFF", "naładowano"), (C0+5, "", AMBER, INK, "naładuj dziś"),
       (C0+9, "", AMBER_L, INK, "zaplanowane"), (C0+13, "", RED_L, INK, "przegapione"), (C0+17, "", "E5E7EB", INK, "święto")]
for col, sym, bgc, fc, txt in leg:
    c = K.cell(5, col, sym); c.fill = fill(bgc); c.font = font(10, True, fc); c.alignment = CENTER; c.border = B_ALL
    t = K.cell(5, col + 1, txt); t.font = font(8, color=MUTED); t.alignment = Alignment(vertical="center")

K.row_dimensions[6].height = 20; K.row_dimensions[7].height = 16; K.row_dimensions[K_HDR].height = 20; K.row_dimensions[K_CNT].height = 18
for r, lab in ((6, ""), (7, ""), (K_HDR, "Samochód"), (K_CNT, "Do naładowania →")):
    paint(K, f"B{r}:D{r}", f=fill(INK))
    K.cell(r, 2, lab).font = font(9, True, "FFFFFF" if r == K_HDR else "9CA3AF"); K.cell(r, 2).alignment = LEFT
K.cell(K_HDR, 3, "VIN (końc.)").font = font(8, True, "FFFFFF"); K.cell(K_HDR, 3).alignment = CENTER
K.cell(K_HDR, 4, "Następne").font = font(8, True, ACCENT); K.cell(K_HDR, 4).alignment = CENTER

SG = f"Samochody!$G${S_FIRST}:$G${S_LAST}"
wk = Side(style="medium", color="9CA3AF")
month_start = None
for j, d in enumerate(DATES):
    col = C0 + j; cl = L(col)
    mfill = fill(INK) if (d.month % 2) else fill(INK2)
    if month_start is None or DATES[j - 1].month != d.month:
        if month_start is not None and month_start < col - 1:
            K.merge_cells(start_row=6, start_column=month_start, end_row=6, end_column=col - 1)
        month_start = col
        K.cell(6, col, f"{MIES[d.month-1].upper()} {d.year}").font = font(9, True, ACCENT)
        K.cell(6, col).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for r in (6, 7, K_HDR, K_CNT): K.cell(r, col).fill = mfill
    K.cell(7, col, ["pn", "wt", "śr", "cz", "pt"][d.weekday()]).font = font(7, color="9CA3AF")
    K.cell(7, col).alignment = CENTER
    hc = K.cell(K_HDR, col, d); hc.number_format = "d"; hc.font = font(9, True, "FFFFFF"); hc.alignment = CENTER
    cc = K.cell(K_CNT, col, f'=IF({cl}${K_HDR}<TODAY(),"",IF({cl}${K_HDR}=TODAY(),COUNTIF({SG},"<="&{cl}${K_HDR}),COUNTIF({SG},{cl}${K_HDR})))')
    cc.number_format = '0;;""'; cc.font = font(8, True, ACCENT); cc.alignment = CENTER
    if d.weekday() == 0:
        for r in range(K_HDR, K_LAST + 1):
            K.cell(r, col).border = Border(left=wk, bottom=thin) if r >= K_FIRST else Border(left=wk)
K.merge_cells(start_row=6, start_column=month_start, end_row=6, end_column=C0 + len(DATES) - 1)

for i in range(NCARS):
    r = K_FIRST + i; sr = S_FIRST + i
    K.row_dimensions[r].height = 24
    K.cell(r, 2, f'=IF(COUNTA(Samochody!C{sr}:E{sr})=0,"",TRIM(Samochody!C{sr}&" "&Samochody!D{sr}))')
    K.cell(r, 3, f'=IF(Samochody!E{sr}="","",RIGHT(Samochody!E{sr},6))')
    K.cell(r, 4, f'=Samochody!G{sr}')
    K.cell(r, 2).font = font(9, True); K.cell(r, 2).alignment = LEFT
    K.cell(r, 3).font = Font(name="Consolas", size=8, color=MUTED); K.cell(r, 3).alignment = CENTER
    K.cell(r, 4).font = font(9, True); K.cell(r, 4).alignment = CENTER; K.cell(r, 4).number_format = "dd.mm"
    for col in (2, 3, 4): K.cell(r, col).border = Border(bottom=thin, right=thin if col == 4 else None)
    for j, d in enumerate(DATES):
        c = K.cell(r, C0 + j)
        c.protection = UNLOCK; c.alignment = CENTER; c.font = font(11, True, "FFFFFF"); c.number_format = '"✔";"✔";;"✔"'
        if d.weekday() != 0: c.border = Border(bottom=thin, right=Side(style="hair", color=LINE))

dvk = DataValidation(type="list", formula1='"x,✔"', allow_blank=True, showErrorMessage=False)
K.add_data_validation(dvk); dvk.add(f"E{K_FIRST}:{CL}{K_LAST}")

G = f"E{K_FIRST}:{CL}{K_LAST}"
e = f"E{K_FIRST}"; h = f"E${K_HDR}"; nx = f"$D{K_FIRST}"; nm = f"$B{K_FIRST}"
rules = [
    (f'AND({e}<>"",{e}<>FALSE,{h}>TODAY())', RED, "FFFFFF"),
    (f'AND({e}<>"",{e}<>FALSE)', GREEN, "FFFFFF"),
    (f'AND({nm}<>"",{nx}<>"",{h}=TODAY(),{nx}<=TODAY())', AMBER, INK),
    (f'AND({nm}<>"",{nx}<>"",{h}<TODAY(),{h}>={nx})', RED_L, INK),
    (f'AND({nm}<>"",{h}={nx})', AMBER_L, INK),
    (f'INDEX(DniRob,{h}-DniStart+1)=0', "E5E7EB", INK),
    (f'{h}=TODAY()', BLUE_L, INK),
]
for f_, bgc, fc in rules:
    K.conditional_formatting.add(G, FormulaRule(formula=[f_], fill=fill(bgc), font=Font(name=FONT, bold=True, color=fc), stopIfTrue=True))
HRNG = f"E6:{CL}{K_CNT}"
K.conditional_formatting.add(f"E7:{CL}{K_CNT}", FormulaRule(formula=["E$8=TODAY()"], fill=fill(ACCENT), font=Font(name=FONT, bold=True, color=INK), stopIfTrue=True))
K.conditional_formatting.add(f"E7:{CL}{K_CNT}", FormulaRule(formula=["INDEX(DniRob,E$8-DniStart+1)=0"], font=Font(name=FONT, italic=True, color="6B7280")))

K.freeze_panes = f"E{K_FIRST}"
today = dt.date.today()
idx = max([j for j, d in enumerate(DATES) if d <= today] or [0])
K.sheet_view.pane.topLeftCell = f"{L(C0 + max(idx - 3, 0))}{K_FIRST}"
K.protection.sheet = True; K.protection.formatColumns = False; K.protection.formatRows = False; K.protection.formatCells = False
K.page_setup.orientation = "landscape"; K.sheet_properties.pageSetUpPr.fitToPage = True
K.page_setup.fitToWidth = 1; K.page_setup.fitToHeight = 0
K.print_title_cols = "B:D"

# =====================================================================
# INSTRUKCJA
# =====================================================================
I.column_dimensions["A"].width = 2.5; I.column_dimensions["B"].width = 6; I.column_dimensions["C"].width = 105
header_band(I, "D", "INSTRUKCJA", "Codzienna obsługa zajmuje ok. 1 minuty.")
steps = [
    ("RAZ NA POCZĄTKU", None),
    ("1", "Zakładka Samochody: wpisz w kolejnych wierszach markę (lista: Dacia / Renault), model i numer VIN. Miejsce na 30 aut."),
    ("2", "Zakładka Ustawienia: wybierz co ile dni ładować (2 lub 3) i sprawdź listę świąt."),
    ("CODZIENNIE (Mateusz)", None),
    ("3", "Otwórz zakładkę Kalendarz i kliknij „PRZEJDŹ DO DZIŚ”. Pomarańczowe kratki w dzisiejszej kolumnie = te auta trzeba dziś naładować."),
    ("4", "Kliknij kratkę w dzisiejszej kolumnie przy pierwszym aucie i idź w dół: przy naładowanym aucie wpisz x i Enter, przy pozostałych samo Enter – kursor schodzi do kolejnego auta."),
    ("5", "Kratka zrobi się zielona, a termin następnego ładowania (kolumna „Następne”) przesunie się sam o 2 dni robocze."),
    ("6", "Pomyłka? Zaznacz kratkę i naciśnij Delete."),
    ("DOBRZE WIEDZIEĆ", None),
    ("•", "Ładowanie w czwartek → następne w piątek. Ładowanie w piątek → następne w poniedziałek. Święta są omijane tak samo."),
    ("•", "Czerwone kratki = dni, w których auto było do naładowania, a nie zostało naładowane."),
    ("•", "Ciemnoczerwona kratka z ✔ = zaznaczenie w przyszłym dniu (pomyłka) – nie jest liczone. Usuń je klawiszem Delete."),
    ("•", "Kalendarz obejmuje dni robocze od 02.11.2026 do 29.12.2028. Weekendów w nim nie ma – ładujemy tylko pn–pt."),
    ("•", "Arkusze są zabezpieczone przed przypadkowym skasowaniem wzorów (bez hasła: Recenzja → Nie chroń arkusza)."),
    ("LIBREOFFICE / EXCEL", None),
    ("•", "Plik działa tak samo w Excelu i w LibreOffice. Zapisuj go zawsze jako .xlsx – w LibreOffice przy zapisie wybierz „Użyj formatu Excel 2007-365!”."),
    ("•", "Plik na wspólnym dysku: w danej chwili edytuje go jedna osoba. Pozostali otwierają go „tylko do odczytu” albo po zamknięciu przez Mateusza."),
]
r = 5
for a, b in steps:
    if b is None:
        r += 1; I.cell(r, 2, a).font = font(10, True); paint(I, f"B{r}:C{r}", border=B_BOT)
    else:
        I.cell(r, 2, a).font = font(11, True, AMBER); I.cell(r, 2).alignment = Alignment(horizontal="center", vertical="top")
        I.cell(r, 3, b).font = font(10); I.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="top")
        I.row_dimensions[r].height = 30
    r += 1

wb.active = 0
wb.save(OUT)
print("zapisano", OUT, "kolumn dat:", len(DATES))
