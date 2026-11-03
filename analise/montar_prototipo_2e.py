"""Prototipo 2.0 v4 - quatro SBs, chaves por formula, feriados."""
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
COD = {"1ªSB": 1, "2ªSB": 2, "3ªSB": 3, "4ªSB": 4, "ADM": 0}

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

ws = wb.active
ws.title = "PARAMETROS"
ws["A1"] = "Competência"
ws["B1"] = datetime.datetime(2026, 11, 1)
ws["B1"].number_format = "dd/mm/yyyy"
for c, v in [("A3", "SB"), ("B3", "Meta BM/dia"), ("C3", "Meta motoristas/dia"), ("D3", "Código")]:
    ws[c] = v
    ws[c].font = Font(bold=True)
for i, (sb, bm, mot) in enumerate([("1ªSB", 8, 2), ("2ªSB", 7, 1), ("3ªSB", 4, 1), ("4ªSB", 7, 1)]):
    ws.cell(4 + i, 1, sb)
    ws.cell(4 + i, 2, bm)
    ws.cell(4 + i, 3, mot)
    ws.cell(4 + i, 4, i + 1)
for col, w in zip("ABCD", (22, 16, 22, 10)):
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

# MES: posicao e chave por formula
ws = wb.create_sheet("MES")
ws.append(["ID", "SEÇÃO", "CÓDIGO", "ESCALA", "ADM", "MOTORISTA", "ORDEM", "POSIÇÃO", "CHAVE"])
for c in ws[1]:
    c.font = hdr
    c.fill = fill
for i, m in enumerate(mes, start=2):
    ws.cell(i, 1, m[0])
    ws.cell(i, 2, m[1])
    ws.cell(i, 3).value = f'=IF($B{i}="","",IFERROR(VLOOKUP($B{i},PARAMETROS!$A$4:$D$7,4,0),9))'
    ws.cell(i, 4, m[2])
    ws.cell(i, 5, m[3])
    ws.cell(i, 6, m[4])
    ws.cell(i, 7, m[5])
    ws.cell(i, 8).value = f'=IF($B{i}="","",COUNTIFS($B$2:$B$200,$B{i},$G$2:$G$200,"<"&$G{i})+1)'
    ws.cell(i, 9).value = f'=IF($B{i}="","",$C{i}&"#"&$H{i})'
ws.freeze_panes = "A2"
for col, w in zip("ABCDEFGHI", (14, 10, 9, 12, 8, 12, 8, 10, 10)):
    ws.column_dimensions[col].width = w

ws = wb.create_sheet("FERIADOS")
ws.append(["DATA", "NOME"])
for c in ws[1]:
    c.font = hdr
    c.fill = fill
feriados = [
    (datetime.datetime(2026, 1, 1), "Confraternizacao Universal"),
    (datetime.datetime(2026, 2, 16), "Carnaval"), (datetime.datetime(2026, 2, 17), "Carnaval"),
    (datetime.datetime(2026, 3, 6), "Data Magna de PE"),
    (datetime.datetime(2026, 4, 3), "Sexta-feira da Paixao"),
    (datetime.datetime(2026, 4, 21), "Tiradentes"),
    (datetime.datetime(2026, 5, 1), "Dia do Trabalho"),
    (datetime.datetime(2026, 6, 4), "Corpus Christi"),
    (datetime.datetime(2026, 6, 24), "Sao Joao (PE)"),
    (datetime.datetime(2026, 6, 29), "Sao Pedro (PE)"),
    (datetime.datetime(2026, 9, 7), "Independencia"),
    (datetime.datetime(2026, 10, 12), "N. S. Aparecida"),
    (datetime.datetime(2026, 11, 2), "Finados"),
    (datetime.datetime(2026, 11, 15), "Proclamacao da Republica"),
    (datetime.datetime(2026, 11, 20), "Consciencia Negra"),
    (datetime.datetime(2026, 12, 25), "Natal"),
]
for d, n in feriados:
    ws.append([d, n])
    ws.cell(ws.max_row, 1).number_format = "dd/mm/yyyy"
ws.column_dimensions["A"].width = 13
ws.column_dimensions["B"].width = 32

N = 42
def montar_grade(nome_aba, codigo_sb, meta_row):
    ws = wb.create_sheet(nome_aba)
    ws["A1"] = f"{nome_aba.replace('GRADE ', '')} — 7º GB — CBMPE"
    ws["A1"].font = Font(bold=True, size=14)
    ws["B1"] = '=TEXT(PARAMETROS!$B$1,"mmmm/yyyy")'
    ws["B1"].font = Font(bold=True, size=12)
    for col, v in {1: "Nº", 2: "ESCALA", 3: "MOT.", 4: "MILITAR", 5: "PJES", 6: "FEM"}.items():
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
        ws.column_dimensions[col].width = 3.8

    for i in range(N):
        r = 3 + i
        ws[f"AN{r}"] = f'=IFERROR(MATCH({codigo_sb}&"#{i + 1}",MES!$I$2:$I$200,0),"")'
        ws[f"A{r}"] = f'=IF($AN{r}="","",{i + 1})'
        ws[f"D{r}"] = (f'=IF($AN{r}="","",IFERROR(INDEX(PESSOAS!$E$2:$E$200,'
                       f'MATCH(INDEX(MES!$A$2:$A$200,$AN{r}),PESSOAS!$A$2:$A$200,0)),""))')
        ws[f"B{r}"] = f'=IF($AN{r}="","",INDEX(MES!$D$2:$D$200,$AN{r}))'
        ws[f"C{r}"] = f'=IF($AN{r}="","",INDEX(MES!$F$2:$F$200,$AN{r}))'

    r_bm, r_mot, r_def = 3 + N + 2, 3 + N + 3, 3 + N + 4
    ws[f"D{r_bm}"] = "MILITARES POR DIA"
    ws[f"D{r_mot}"] = "MOTORISTAS POR DIA"
    ws[f"D{r_def}"] = "DÉFICIT / SOBRA"
    for r in (r_bm, r_mot, r_def):
        ws[f"D{r}"].font = Font(bold=True)

    for i in range(31):
        col = get_column_letter(7 + i)
        base = f"{col}$3:{col}${2 + N}"
        cbase = f"$C$3:$C${2 + N}"
        ws[f"{col}{r_bm}"] = (f'=IF({col}$2="","",COUNTIF({base},"x")+COUNTIF({base},"PJ")+COUNTIF({base},"SPJ")'
                              f'+(COUNTIF({base},"PJ12")+COUNTIF({base},"PJ12H")+COUNTIF({base},"S1"))*0.5)')
        ws[f"{col}{r_mot}"] = (f'=IF({col}$2="","",COUNTIFS({cbase},"M",{base},"x")+COUNTIFS({cbase},"M",{base},"PJ")'
                               f'+COUNTIFS({cbase},"M",{base},"SPJ")'
                               f'+(COUNTIFS({cbase},"M",{base},"PJ12")+COUNTIFS({cbase},"M",{base},"PJ12H")'
                               f'+COUNTIFS({cbase},"M",{base},"S1"))*0.5)')
        ws[f"{col}{r_def}"] = f'=IF({col}$2="","",PARAMETROS!$B${meta_row}-{col}{r_bm})'
        ws[f"{col}45"] = (f'=IF({col}$2="","",IF(OR(WEEKDAY({col}$2,2)>5,'
                          f'COUNTIF(FERIADOS!$A:$A,{col}$2)>0),1,0))')
    ws["D45"] = "aux: fim de semana/feriado"
    ws.row_dimensions[45].hidden = True

    ws.conditional_formatting.add(f"G3:AK{2 + N}", FormulaRule(
        formula=['G$45=1'], fill=PatternFill("solid", fgColor="FFF4CCCC")))
    ws.conditional_formatting.add("G2:AK2", FormulaRule(
        formula=['G$45=1'], fill=PatternFill("solid", fgColor="FFEA9999")))

    for col, w in zip(["A", "B", "C", "D", "E", "F"], (4, 9, 6, 30, 6, 5)):
        ws.column_dimensions[col].width = w
    ws.column_dimensions["AN"].hidden = True
    ws.freeze_panes = "G3"

for nome, cod, meta in [("GRADE 1ªSB", 1, 4), ("GRADE 2ªSB", 2, 5),
                        ("GRADE 3ªSB", 3, 6), ("GRADE 4ªSB", 4, 7)]:
    montar_grade(nome, cod, meta)

wb.save(OUT)
print("salvo:", OUT)
for n, c, _ in [("GRADE 1ªSB", 1, 4), ("GRADE 2ªSB", 2, 5), ("GRADE 3ªSB", 3, 6), ("GRADE 4ªSB", 4, 7)]:
    q = sum(1 for m in mes if COD.get(m[1]) == c)
    print(f"  {n}: {q} militares")
