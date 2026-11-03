"""Prototipo do 2.0: PESSOAS, MES, PARAMETROS e a GRADE da 4aSB derivada."""
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

ARQS = {
    "AGO": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\8 - Escala de Agosto 2026.xlsx",
    "SET": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\9 - Escala de Setembro 2026.xlsx",
    "OUT": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx",
    "NOV": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\11 - Escala de Novembro 2026.xlsx",
}
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_2.0_prototipo.xlsx"
from openpyxl import load_workbook

# ---------- dados: pessoas (uniao) e mes de novembro ----------
pessoas = {}
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

# ---------- workbook ----------
wb = Workbook()
hdr = Font(bold=True, color="FFFFFF")
fill = PatternFill("solid", fgColor="4A6FA5")

# PARAMETROS
ws = wb.active
ws.title = "PARAMETROS"
ws["A1"] = "Competência"
ws["B1"] = datetime.datetime(2026, 11, 1)
ws["B1"].number_format = "dd/mm/yyyy"
ws["A3"], ws["B3"], ws["C3"] = "SB", "Meta BM/dia", "Meta motoristas/dia"
for c in ws[3]:
    c.font = Font(bold=True)
for i, (sb, bm, mot) in enumerate([("1ªSB", 8, 2), ("2ªSB", 7, 1), ("3ªSB", 4, 1), ("4ªSB", 7, 1)]):
    ws.cell(4 + i, 1).value = sb
    ws.cell(4 + i, 2).value = bm
    ws.cell(4 + i, 3).value = mot
ws["A10"] = "Pesos"
ws["A11"], ws["B11"] = "X / PJ / SPJ", 1
ws["A12"], ws["B12"] = "PJ12 / PJ12H / S1", 0.5
ws["A13"], ws["B13"] = "SB / N", 0
ws.column_dimensions["A"].width = 22
ws.column_dimensions["B"].width = 16
ws.column_dimensions["C"].width = 20

# PESSOAS
ws = wb.create_sheet("PESSOAS")
ws.append(["ID", "POSTO", "NOME DE GUERRA", "ORDEM", "NOME COMPLETO"])
for c in ws[1]:
    c.font = hdr
    c.fill = fill
for mat, (posto, guerra) in pessoas.items():
    r = ws.max_row + 1
    ws.append([mat, posto, guerra, None, f'=IF(B{r}="","",B{r}&" "&A{r}&" "&C{r})'])
# ordem = a do mes de novembro quando existir, senao alfabetica
ordem_de = {m[0]: m[5] for m in mes}
faltando = sorted([m for m in pessoas if m not in ordem_de])
prox = len(ordem_de)
for r in range(2, ws.max_row + 1):
    mat = ws.cell(r, 1).value
    if mat in ordem_de:
        ws.cell(r, 4).value = ordem_de[mat]
    else:
        prox += 1
        ws.cell(r, 4).value = prox
ws.freeze_panes = "A2"
for col, w in zip("ABCDE", (14, 10, 24, 8, 34)):
    ws.column_dimensions[col].width = w

# MES
ws = wb.create_sheet("MES")
ws.append(["ID", "SEÇÃO", "ESCALA", "ADM", "MOTORISTA"])
for c in ws[1]:
    c.font = hdr
    c.fill = fill
for m in mes:
    ws.append([m[0], m[1], m[2], m[3], m[4]])
ws.freeze_panes = "A2"
for col, w in zip("ABCDE", (14, 10, 12, 8, 12)):
    ws.column_dimensions[col].width = w

# GRADE 4aSB
ws = wb.create_sheet("GRADE 4ªSB")
ws["A1"] = "4ªSB — 7º GB — CBMPE"
ws["A1"].font = Font(bold=True, size=14)
ws["B1"] = '=TEXT(PARAMETROS!$B$1,"mmmm/yyyy")'
ws["B1"].font = Font(bold=True, size=12)

# bloco auxiliar (oculto) com o efetivo derivado
AUX = 40   # coluna AN
cols = ["AN", "AO", "AP", "AQ"]
ws[f"{cols[0]}1"] = "aux: nome"
ws[f"{cols[1]}1"] = "aux: escala"
ws[f"{cols[2]}1"] = "aux: adm"
ws[f"{cols[3]}1"] = "aux: motorista"
sel = 'MES!$B$2:$B$200="4ªSB"'
ws[f"{cols[0]}2"] = (
    f'=IFERROR(SORTBY(FILTER({{PESSOAS!$E$2:$E$200,PESSOAS!$D$2:$D$200}},{sel}),'
    f'FILTER(PESSOAS!$D$2:$D$200,{sel})),'
    f'SORTBY(FILTER({{PESSOAS!$E$2:$E$200,PESSOAS!$D$2:$D$200}},{sel}),'
    f'FILTER(PESSOAS!$D$2:$D$200,{sel})))'
)

# cabecalho de dias
ws["A2"], ws["B2"], ws["C2"], ws["D2"], ws["E2"], ws["F2"] = "Nº", "ESCALA", "MOT.", "MILITAR", "PJES", "FEM"
for c in ws[2]:
    if c.value:
        c.font = hdr
        c.fill = fill
        c.alignment = Alignment(horizontal="center")
for i in range(31):
    col = get_column_letter(7 + i)
    ws[f"{col}2"] = (f"=IF(MONTH(PARAMETROS!$B$1+{i})=MONTH(PARAMETROS!$B$1),"
                     f"PARAMETROS!$B$1+{i},\"\")")
    ws[f"{col}2"].number_format = "dd"
    ws[f"{col}2"].font = hdr
    ws[f"{col}2"].fill = fill
    ws[f"{col}2"].alignment = Alignment(horizontal="center")
    ws.column_dimensions[col].width = 3.6

# linhas de militares
for i in range(40):
    r = 3 + i
    ws[f"A{r}"] = f'=IF($AN{r}="","",{i + 1})'
    ws[f"D{r}"] = f'=IFERROR(INDEX($AN$2:$AO$60,{i + 1},1),"")'
    ws[f"B{r}"] = f'=IF($D{r}="","",IFERROR(INDEX(MES!$C$2:$C$200,MATCH($D{r},PESSOAS!$E$2:$E$200,0)),""))'
    ws[f"C{r}"] = f'=IF($D{r}="","",IFERROR(INDEX(MES!$E$2:$E$200,MATCH($D{r},PESSOAS!$E$2:$E$200,0)),""))'

# totais
r_tot = 47
ws[f"D{r_tot}"] = "MILITARES POR DIA"
ws[f"D{r_tot}"].font = Font(bold=True)
ws[f"D{r_tot + 1}"] = "MOTORISTAS POR DIA"
ws[f"D{r_tot + 1}"].font = Font(bold=True)
meta_bm = 'PARAMETROS!$B$7'
meta_mot = 'PARAMETROS!$C$7'
for i in range(31):
    col = get_column_letter(7 + i)
    ws[f"{col}{r_tot}"] = (
        f'=IF({col}$2="","",SUMPRODUCT(COUNTIF({col}$3:{col}42,{{"x";"PJ";"SPJ"}}))'
        f'+(SUMPRODUCT(COUNTIF({col}$3:{col}42,{{"PJ12";"PJ12H";"S1"}}))*0.5))'
    )
    ws[f"{col}{r_tot + 1}"] = (
        f'=IF({col}$2="","",SUMPRODUCT(COUNTIFS($C$3:$C$42,"M",{col}$3:{col}42,{{"x";"PJ";"SPJ"}}))'
        f'+(SUMPRODUCT(COUNTIFS($C$3:$C$42,"M",{col}$3:{col}42,{{"PJ12";"PJ12H";"S1"}}))*0.5))'
    )
    ws[f"{col}{r_tot + 2}"] = f'=IF({col}$2="","",({meta_bm}-{col}{r_tot})*-1)'
    ws.column_dimensions[col].width = 3.6
ws[f"D{r_tot + 2}"] = "DÉFICIT/SOBRA"
ws.column_dimensions["D"].width = 30
ws.column_dimensions["A"].width = 4
ws.column_dimensions["B"].width = 9
ws.column_dimensions["C"].width = 6
ws.column_dimensions["E"].width = 6
ws.column_dimensions["F"].width = 5
ws.freeze_panes = "G3"

for c in ["AN", "AO", "AP", "AQ"]:
    ws.column_dimensions[c].hidden = True

wb.save(OUT)
print("salvo:", OUT)
print("pessoas:", len(pessoas), "| linhas no MES:", len(mes))
