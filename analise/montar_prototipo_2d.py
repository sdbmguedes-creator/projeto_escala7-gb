"""Prototipo 2.0 v3 - sem FILTER/SORTBY e sem chaves de matriz."""
import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import FormulaRule
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

# posicao dentro da secao (ordem de antiguidade) - chave em ASCII puro
ASCII = {"1ªSB": "SB1", "2ªSB": "SB2", "3ªSB": "SB3", "4ªSB": "SB4"}
por_secao = {}
for m in sorted(mes, key=lambda x: x[5]):
    por_secao.setdefault(m[1], []).append(m[0])
pos = {}
for sb, lst in por_secao.items():
    for i, mat in enumerate(lst, 1):
        pos[mat] = (i, f"{ASCII.get(sb, sb)}#{i}")

wb = Workbook()
hdr = Font(bold=True, color="FFFFFF")
fill = PatternFill("solid", fgColor="4A6FA5")

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

ws = wb.create_sheet("PESSOAS")
ws.append(["ID", "POSTO", "NOME DE GUERRA", "ORDEM", "NOME COMPLETO"])
for c in ws[1]:
    c.font = hdr
    c.fill = fill
for mat, (posto, guerra) in pessoas.items():
    r = ws.max_row + 1
    ws.append([mat, posto, guerra, ordem_de.get(mat, 999),
               f'=IF(B{r}="","",B{r}&" "&A{r}&" "&C{r})'])
ws.freeze_panes = "A2"
for col, w in zip("ABCDE", (14, 10, 24, 8, 34)):
    ws.column_dimensions[col].width = w

ws = wb.create_sheet("MES")
ws.append(["ID", "SEÇÃO", "ESCALA", "ADM", "MOTORISTA", "ORDEM", "POSIÇÃO", "CHAVE"])
for c in ws[1]:
    c.font = hdr
    c.fill = fill
for m in mes:
    r = ws.max_row + 1
    p, chave = pos[m[0]]
    ws.append([m[0], m[1], m[2], m[3], m[4], m[5], p, chave])
ws.freeze_panes = "A2"
for col, w in zip("ABCDEFGH", (14, 10, 12, 8, 12, 8, 10, 14)):
    ws.column_dimensions[col].width = w

ws = wb.create_sheet("FERIADOS")
ws.append(["DATA", "NOME"])
for c in ws[1]:
    c.font = hdr
    c.fill = fill
feriados = [
    (datetime.datetime(2026, 1, 1), "Confraternização Universal"),
    (datetime.datetime(2026, 2, 16), "Carnaval"),
    (datetime.datetime(2026, 2, 17), "Carnaval"),
    (datetime.datetime(2026, 3, 6), "Data Magna de PE"),
    (datetime.datetime(2026, 4, 3), "Sexta-feira da Paixão"),
    (datetime.datetime(2026, 4, 21), "Tiradentes"),
    (datetime.datetime(2026, 5, 1), "Dia do Trabalho"),
    (datetime.datetime(2026, 6, 4), "Corpus Christi"),
    (datetime.datetime(2026, 6, 24), "São João (PE)"),
    (datetime.datetime(2026, 6, 29), "São Pedro (PE)"),
    (datetime.datetime(2026, 9, 7), "Independência"),
    (datetime.datetime(2026, 10, 12), "Nossa Senhora Aparecida"),
    (datetime.datetime(2026, 11, 2), "Finados"),
    (datetime.datetime(2026, 11, 15), "Proclamação da República"),
    (datetime.datetime(2026, 11, 20), "Consciência Negra"),
    (datetime.datetime(2026, 12, 25), "Natal"),
]
for d, n in feriados:
    ws.append([d, n])
    ws.cell(ws.max_row, 1).number_format = "dd/mm/yyyy"
ws.column_dimensions["A"].width = 13
ws.column_dimensions["B"].width = 30

ws = wb.create_sheet("GRADE 4ªSB")
ws["A1"] = "4ªSB — 7º GB — CBMPE"
ws["A1"].font = Font(bold=True, size=14)
ws["B1"] = '=TEXT(PARAMETROS!$B$1,"mmmm/yyyy")'
ws["B1"].font = Font(bold=True, size=12)

cab = {1: "Nº", 2: "ESCALA", 3: "MOT.", 4: "MILITAR", 5: "PJES", 6: "FEM"}
for col, v in cab.items():
    c = ws.cell(2, col, v)
    c.font = hdr
    c.fill = fill
    c.alignment = Alignment(horizontal="center")
for i in range(31):
    col = get_column_letter(7 + i)
    c = ws[f"{col}2"]
    c.value = (f'=IF(MONTH(PARAMETROS!$B$1+{i})=MONTH(PARAMETROS!$B$1),'
               f'PARAMETROS!$B$1+{i},"")')
    c.number_format = "dd"
    c.font = hdr
    c.fill = fill
    c.alignment = Alignment(horizontal="center")

N = 42
for i in range(N):
    r = 3 + i
    ws[f"AN{r}"] = f'=IFERROR(MATCH("SB4#{i + 1}",MES!$H$2:$H$200,0),"")'
    ws[f"A{r}"] = f'=IF($AN{r}="","",{i + 1})'
    ws[f"D{r}"] = (f'=IF($AN{r}="","",IFERROR(INDEX(PESSOAS!$E$2:$E$200,'
                   f'MATCH(INDEX(MES!$A$2:$A$200,$AN{r}),PESSOAS!$A$2:$A$200,0)),""))')
    ws[f"B{r}"] = f'=IF($AN{r}="","",INDEX(MES!$C$2:$C$200,$AN{r}))'
    ws[f"C{r}"] = f'=IF($AN{r}="","",INDEX(MES!$E$2:$E$200,$AN{r}))'

r_bm, r_mot, r_def = 3 + N + 2, 3 + N + 3, 3 + N + 4
ws[f"D{r_bm}"] = "MILITARES POR DIA"
ws[f"D{r_mot}"] = "MOTORISTAS POR DIA"
ws[f"D{r_def}"] = "DÉFICIT / SOBRA"
for r in (r_bm, r_mot, r_def):
    ws[f"D{r}"].font = Font(bold=True)

faixa = f"$3:${2 + N}"
for i in range(31):
    col = get_column_letter(7 + i)
    base = f"{col}{faixa}"
    ws[f"{col}{r_bm}"] = (f'=IF({col}$2="","",COUNTIF({base},"x")+COUNTIF({base},"PJ")+COUNTIF({base},"SPJ")'
                          f'+(COUNTIF({base},"PJ12")+COUNTIF({base},"PJ12H")+COUNTIF({base},"S1"))*0.5)')
    ws[f"{col}{r_mot}"] = (f'=IF({col}$2="","",COUNTIFS($C{faixa},"M",{base},"x")+COUNTIFS($C{faixa},"M",{base},"PJ")+COUNTIFS($C{faixa},"M",{base},"SPJ")'
                           f'+(COUNTIFS($C{faixa},"M",{base},"PJ12")+COUNTIFS($C{faixa},"M",{base},"PJ12H")+COUNTIFS($C{faixa},"M",{base},"S1"))*0.5)')
    ws[f"{col}{r_def}"] = f'=IF({col}$2="","",PARAMETROS!$B$7-{col}{r_bm})'
    ws.column_dimensions[col].width = 3.8

for col, w in zip(["A", "B", "C", "D", "E", "F"], (4, 9, 6, 30, 6, 5)):
    ws.column_dimensions[col].width = w
ws.column_dimensions["AN"].hidden = True
ws.freeze_panes = "G3"

# linha auxiliar 45: marca dias de fim de semana / feriado (CF nao le outra aba)
ws["D45"] = "aux: fim de semana/feriado"
for i in range(31):
    col = get_column_letter(7 + i)
    ws[f"{col}45"] = (f'=IF({col}$2="","",IF(OR(WEEKDAY({col}$2,2)>5,'
                      f'COUNTIF(FERIADOS!$A:$A,{col}$2)>0),1,0))')
ws.row_dimensions[45].hidden = True

# fim de semana / feriado em vermelho (le a linha 45, mesma aba)
area = f"G3:AK{2 + N}"
ws.conditional_formatting.add(area, FormulaRule(
    formula=['G$45=1'],
    fill=PatternFill("solid", fgColor="FFF4CCCC")))
ws.conditional_formatting.add("G2:AK2", FormulaRule(
    formula=['G$45=1'],
    fill=PatternFill("solid", fgColor="FFEA9999")))

wb.save(OUT)
print("salvo:", OUT)
print("4ªSB:", sum(1 for m in mes if m[1] == "4ªSB"))
print("exemplo CHAVE:", mes[0][0], "->", pos[mes[0][0]][1])
