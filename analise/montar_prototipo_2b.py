"""Prototipo 2.0 - grade derivada, com juncao correta por ID."""
import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

ARQS = {
    "AGO": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\8 - Escala de Agosto 2026.xlsx",
    "SET": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\9 - Escala de Setembro 2026.xlsx",
    "OUT": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx",
    "NOV": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\11 - Escala de Novembro 2026.xlsx",
}
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_2.0_prototipo.xlsx"

pessoas, ordem_de = {}, {}
for p in ARQS.values():
    wv = load_workbook(p, data_only=True)
    res = wv["RESUMO"]
    for r in range(2, 160):
        posto, mat, guerra = (res.cell(r, c).value for c in (2, 3, 4))
        if all(isinstance(x, str) for x in (posto, mat, guerra)):
            pessoas.setdefault(mat.strip(), (posto.strip(), guerra.strip()))

wv = load_workbook(ARQS["NOV"], data_only=True)
res = wv["RESUMO"]
mes, vistos, ordem = [], set(), 0
for r in range(2, 160):
    posto, mat, guerra = (res.cell(r, c).value for c in (2, 3, 4))
    if not all(isinstance(x, str) for x in (posto, mat, guerra)):
        continue
    mat = mat.strip()
    if mat in vistos:
        continue
    vistos.add(mat)
    ordem += 1
    esc = res.cell(r, 8).value
    mes.append((mat, res.cell(r, 5).value, esc,
                "S" if isinstance(esc, str) and "ADM" in esc else "",
                res.cell(r, 7).value, ordem))
    ordem_de[mat] = ordem

wb = Workbook()
hdr = Font(bold=True, color="FFFFFF")
fill = PatternFill("solid", fgColor="4A6FA5")

# --- PARAMETROS ---
ws = wb.active
ws.title = "PARAMETROS"
ws["A1"] = "Competência"
ws["B1"] = datetime.datetime(2026, 11, 1)
ws["B1"].number_format = "dd/mm/yyyy"
for c, v in [("A3", "SB"), ("B3", "Meta BM/dia"), ("C3", "Meta motoristas/dia")]:
    ws[c] = v
    ws[c].font = Font(bold=True)
for i, (sb, bm, mot) in enumerate([("1ªSB", 8, 2), ("2ªSB", 7, 1), ("3ªSB", 4, 1), ("4ªSB", 7, 1)]):
    ws.cell(4 + i, 1, sb)
    ws.cell(4 + i, 2, bm)
    ws.cell(4 + i, 3, mot)
for col, w in zip("ABC", (22, 16, 22)):
    ws.column_dimensions[col].width = w

# --- PESSOAS ---
ws = wb.create_sheet("PESSOAS")
ws.append(["ID", "POSTO", "NOME DE GUERRA", "ORDEM", "NOME COMPLETO"])
for c in ws[1]:
    c.font = hdr
    c.fill = fill
prox = max(ordem_de.values()) if ordem_de else 0
for mat, (posto, guerra) in pessoas.items():
    r = ws.max_row + 1
    o = ordem_de.get(mat)
    if o is None:
        prox += 1
        o = prox
    ws.append([mat, posto, guerra, o, f'=IF(B{r}="","",B{r}&" "&A{r}&" "&C{r})'])
ws.freeze_panes = "A2"
for col, w in zip("ABCDE", (14, 10, 24, 8, 34)):
    ws.column_dimensions[col].width = w

# --- MES ---
ws = wb.create_sheet("MES")
ws.append(["ID", "SEÇÃO", "ESCALA", "ADM", "MOTORISTA", "ORDEM"])
for c in ws[1]:
    c.font = hdr
    c.fill = fill
for m in mes:
    ws.append(list(m))
ws.freeze_panes = "A2"
for col, w in zip("ABCDEF", (14, 10, 12, 8, 12, 8)):
    ws.column_dimensions[col].width = w

# --- GRADE 4aSB ---
ws = wb.create_sheet("GRADE 4ªSB")
ws["A1"] = "4ªSB — 7º GB — CBMPE"
ws["A1"].font = Font(bold=True, size=14)
ws["B1"] = '=TEXT(PARAMETROS!$B$1,"mmmm/yyyy")'
ws["B1"].font = Font(bold=True, size=12)

# auxiliar: IDs da 4aSB ordenados pela antiguidade
sel = 'MES!$B$2:$B$200="4ªSB"'
ws["AN1"] = "aux: ids"
ws["AN2"] = f'=IFERROR(SORTBY(FILTER(MES!$A$2:$A$200,{sel}),FILTER(MES!$F$2:$F$200,{sel})),"")'

cab = {1: "Nº", 2: "ESCALA", 3: "MOT.", 4: "MILITAR", 5: "PJES", 6: "FEM"}
for col, v in cab.items():
    c = ws.cell(2, col, v)
    c.font = hdr
    c.fill = fill
    c.alignment = Alignment(horizontal="center")
for i in range(31):
    col = get_column_letter(7 + i)
    c = ws[f"{col}2"]
    c.value = (f"=IF(MONTH(PARAMETROS!$B$1+{i})=MONTH(PARAMETROS!$B$1),"
               f"PARAMETROS!$B$1+{i},\"\")")
    c.number_format = "dd"
    c.font = hdr
    c.fill = fill
    c.alignment = Alignment(horizontal="center")

N = 42
for i in range(N):
    r = 3 + i
    ws[f"A{r}"] = f'=IF($AN{r}="","",{i + 1})'
    ws[f"D{r}"] = (f'=IFERROR(XLOOKUP(INDEX($AN$2:$AN$60,{i + 1}),'
                   f'PESSOAS!$A$2:$A$200,PESSOAS!$E$2:$E$200,""),"")')
    ws[f"B{r}"] = (f'=IF($D{r}="","",IFERROR(XLOOKUP(INDEX($AN$2:$AN$60,{i + 1}),'
                   f'MES!$A$2:$A$200,MES!$C$2:$C$200,""),""))')
    ws[f"C{r}"] = (f'=IF($D{r}="","",IFERROR(XLOOKUP(INDEX($AN$2:$AN$60,{i + 1}),'
                   f'MES!$A$2:$A$200,MES!$E$2:$E$200,""),""))')

r_bm, r_mot, r_def = 3 + N + 2, 3 + N + 3, 3 + N + 4
ws[f"D{r_bm}"] = "MILITARES POR DIA"
ws[f"D{r_mot}"] = "MOTORISTAS POR DIA"
ws[f"D{r_def}"] = "DÉFICIT / SOBRA"
for r in (r_bm, r_mot, r_def):
    ws[f"D{r}"].font = Font(bold=True)

faixa = f"$3:${2 + N}"
for i in range(31):
    col = get_column_letter(7 + i)
    ws[f"{col}{r_bm}"] = (f'=IF({col}$2="","",SUMPRODUCT(COUNTIF({col}{faixa},{{"x";"PJ";"SPJ"}}))'
                          f'+(SUMPRODUCT(COUNTIF({col}{faixa},{{"PJ12";"PJ12H";"S1"}}))*0.5))')
    ws[f"{col}{r_mot}"] = (f'=IF({col}$2="","",SUMPRODUCT(COUNTIFS($C{faixa},"M",{col}{faixa},{{"x";"PJ";"SPJ"}}))'
                           f'+(SUMPRODUCT(COUNTIFS($C{faixa},"M",{col}{faixa},{{"PJ12";"PJ12H";"S1"}}))*0.5))')
    ws[f"{col}{r_def}"] = f'=IF({col}$2="","",PARAMETROS!$B$7-{col}{r_bm}'
    ws.column_dimensions[col].width = 3.8

for col, w in zip(["A", "B", "C", "D", "E", "F"], (4, 9, 6, 30, 6, 5)):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "G3"
ws.column_dimensions["AN"].hidden = True

wb.save(OUT)
print("salvo:", OUT)
print("pessoas:", len(pessoas), "| mes:", len(mes))
print("4aSB no mes:", sum(1 for m in mes if m[1] == "4ªSB"))
