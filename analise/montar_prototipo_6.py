"""Prototipo 2.0 v6 - adiciona GRADE ADM, DIARIO e PJES."""
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
ws["A8"] = "Cotas PJES do mês (SEI)"
ws["A9"], ws["B9"] = "1ªSB", None
ws["A10"], ws["B10"] = "2ªSB", None
ws["A11"], ws["B11"] = "3ªSB", None
ws["A12"], ws["B12"] = "4ªSB", None
ws["A13"], ws["B13"] = "ADM", None
ws["A14"], ws["B14"] = "TOTAL", "=SUM(B9:B13)"
for col, w in zip("ABCD", (26, 16, 22, 10)):
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
ws.append(["ID", "SEÇÃO", "CÓDIGO", "ESCALA", "ADM", "MOTORISTA", "ORDEM", "POSIÇÃO", "CHAVE"])
for c in ws[1]:
    c.font = hdr
    c.fill = fill
for i, m in enumerate(mes, start=2):
    ws.cell(i, 1, m[0])
    ws.cell(i, 2, m[1])
    ws.cell(i, 3).value = f'=IF($B{i}="","",IFERROR(VLOOKUP($B{i},PARAMETROS!$A$4:$D$7,4,0),0))'
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
for d, n in [
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
]:
    ws.append([d, n])
    ws.cell(ws.max_row, 1).number_format = "dd/mm/yyyy"
ws.column_dimensions["A"].width = 13
ws.column_dimensions["B"].width = 32

N = 42
R_BM = 3 + N + 2

def montar_grade(nome_exib, titulo, codigo, meta_row):
    ws = wb.create_sheet(nome_exib)
    ws["A1"] = titulo
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
        ws[f"AN{r}"] = f'=IFERROR(MATCH({codigo}&"#{i + 1}",MES!$I$2:$I$200,0),"")'
        ws[f"A{r}"] = f'=IF($AN{r}="","",{i + 1})'
        ws[f"D{r}"] = (f'=IF($AN{r}="","",IFERROR(INDEX(PESSOAS!$E$2:$E$200,'
                       f'MATCH(INDEX(MES!$A$2:$A$200,$AN{r}),PESSOAS!$A$2:$A$200,0)),""))')
        ws[f"B{r}"] = f'=IF($AN{r}="","",INDEX(MES!$D$2:$D$200,$AN{r}))'
        ws[f"C{r}"] = f'=IF($AN{r}="","",INDEX(MES!$F$2:$F$200,$AN{r}))'
        faixa_linha = f"G{r}:AK{r}"
        ws[f"E{r}"] = (f'=IF($AN{r}="","",COUNTIF({faixa_linha},"PJ")*2+COUNTIF({faixa_linha},"PJ12")'
                       f'+COUNTIF({faixa_linha},"PJ12H")+COUNTIF({faixa_linha},"SPJ"))')

    r_mot, r_def = R_BM + 1, R_BM + 2
    ws[f"D{R_BM}"] = "MILITARES POR DIA"
    ws[f"D{r_mot}"] = "MOTORISTAS POR DIA"
    ws[f"D{r_def}"] = "DÉFICIT / SOBRA"
    ws["D46"] = "TOTAL DE PJES"
    for r in (R_BM, r_mot, r_def, 46):
        ws[f"D{r}"].font = Font(bold=True)

    for i in range(31):
        col = get_column_letter(7 + i)
        base = f"{col}$3:{col}${2 + N}"
        cbase = f"$C$3:$C${2 + N}"
        ws[f"{col}{R_BM}"] = (f'=IF({col}$2="","",COUNTIF({base},"x")+COUNTIF({base},"PJ")+COUNTIF({base},"SPJ")'
                              f'+(COUNTIF({base},"PJ12")+COUNTIF({base},"PJ12H")+COUNTIF({base},"S1"))*0.5)')
        ws[f"{col}{r_mot}"] = (f'=IF({col}$2="","",COUNTIFS({cbase},"M",{base},"x")+COUNTIFS({cbase},"M",{base},"PJ")'
                               f'+COUNTIFS({cbase},"M",{base},"SPJ")'
                               f'+(COUNTIFS({cbase},"M",{base},"PJ12")+COUNTIFS({cbase},"M",{base},"PJ12H")'
                               f'+COUNTIFS({cbase},"M",{base},"S1"))*0.5)')
        if meta_row:
            ws[f"{col}{r_def}"] = f'=IF({col}$2="","",PARAMETROS!$B${meta_row}-{col}{R_BM})'
        else:
            ws[f"{col}{r_def}"] = '=""'
        ws[f"{col}45"] = (f'=IF({col}$2="","",IF(OR(WEEKDAY({col}$2,2)>5,'
                          f'COUNTIF(FERIADOS!$A:$A,{col}$2)>0),1,0))')
    ws["D45"] = "aux: fim de semana/feriado"
    ws.row_dimensions[45].hidden = True
    ws.row_dimensions[46].hidden = False

    ws.conditional_formatting.add(f"G3:AK{2 + N}", FormulaRule(
        formula=['G$45=1'], fill=PatternFill("solid", fgColor="FFF4CCCC")))
    ws.conditional_formatting.add("G2:AK2", FormulaRule(
        formula=['G$45=1'], fill=PatternFill("solid", fgColor="FFEA9999")))

    for col, w in zip(["A", "B", "C", "D", "E", "F"], (4, 9, 6, 30, 6, 5)):
        ws.column_dimensions[col].width = w
    ws.column_dimensions["AN"].hidden = True
    ws.freeze_panes = "G3"
    return ws

for nome, titulo, codigo, meta in [
    ("GRADE 1ªSB", "1ªSB — 7º GB — CBMPE", 1, 4),
    ("GRADE 2ªSB", "2ªSB — 7º GB — CBMPE", 2, 5),
    ("GRADE 3ªSB", "3ªSB — 7º GB — CBMPE", 3, 6),
    ("GRADE 4ªSB", "4ªSB — 7º GB — CBMPE", 4, 7),
    ("GRADE ADM", "ADM — 7º GB — CBMPE", 0, None),
]:
    montar_grade(nome, titulo, codigo, meta)

# ---------- DIARIO ----------
ws = wb.create_sheet("DIARIO")
ws["A1"] = "DIÁRIO — 7º GB"
ws["A1"].font = Font(bold=True, size=14)
ws["B1"] = '=TEXT(PARAMETROS!$B$1,"mmmm/yyyy")'
ws["A2"] = "Dia"
for i in range(31):
    col = get_column_letter(2 + i)
    c = ws[f"{col}2"]
    c.value = f"=IF('GRADE 1ªSB'!{get_column_letter(7 + i)}$2=\"\",\"\",'GRADE 1ªSB'!{get_column_letter(7 + i)}$2)"
    c.number_format = "dd"
    c.font = hdr
    c.fill = fill
    c.alignment = Alignment(horizontal="center")
    ws.column_dimensions[col].width = 4
ws["A2"].font = hdr
ws["A2"].fill = fill

linhas = [("1ªSB", "GRADE 1ªSB", "B"), ("2ªSB", "GRADE 2ªSB", "B"), ("3ªSB", "GRADE 3ªSB", "B"),
          ("4ªSB", "GRADE 4ªSB", "B"), ("ADM", "GRADE ADM", "B")]
r = 3
for rotulo, aba, _ in linhas:
    ws.cell(r, 1, rotulo).font = Font(bold=True)
    for i in range(31):
        col = get_column_letter(2 + i)
        gcol = get_column_letter(7 + i)
        ws[f"{col}{r}"] = f"=IF({col}$2=\"\",\"\",'{aba}'!{gcol}{R_BM})"
    r += 1
ws.cell(r, 1, "TOTAL BM").font = Font(bold=True)
for i in range(31):
    col = get_column_letter(2 + i)
    ws[f"{col}{r}"] = f"=IF({col}$2=\"\",\"\",SUM({col}3:{col}{r - 1}))"
r += 1
ws.cell(r, 1, "MOTORISTAS").font = Font(bold=True)
for i in range(31):
    col = get_column_letter(2 + i)
    ws[f"{col}{r}"] = f"=IF({col}$2=\"\",\"\",SUM('GRADE 1ªSB'!{get_column_letter(7 + i)}{R_BM + 1},'GRADE 2ªSB'!{get_column_letter(7 + i)}{R_BM + 1},'GRADE 3ªSB'!{get_column_letter(7 + i)}{R_BM + 1},'GRADE 4ªSB'!{get_column_letter(7 + i)}{R_BM + 1}))"
r += 1
ws.cell(r, 1, "META 7ºGB").font = Font(bold=True)
for i in range(31):
    col = get_column_letter(2 + i)
    ws[f"{col}{r}"] = "=SUM(PARAMETROS!$B$4:$B$7)"
r += 1
ws.cell(r, 1, "DÉFICIT").font = Font(bold=True)
for i in range(31):
    col = get_column_letter(2 + i)
    ws[f"{col}{r}"] = f"=IF({col}$2=\"\",\"\",{col}{r - 1}-{col}{r - 3})"
ws.column_dimensions["A"].width = 16
ws.freeze_panes = "B3"

# ---------- PJES ----------
ws = wb.create_sheet("PJES")
ws["A1"] = "PJES — cotas do mês"
ws["A1"].font = Font(bold=True, size=14)
cab = ["SB", "Cotas recebidas (SEI)", "Utilizadas", "Saldo", "Situação"]
for i, v in enumerate(cab, start=1):
    c = ws.cell(3, i, v)
    c.font = hdr
    c.fill = fill
ws["A3"] = "SB"
ws["B3"] = "Cotas recebidas (SEI) e demais"
for i, (sb, aba) in enumerate([("1ªSB", "GRADE 1ªSB"), ("2ªSB", "GRADE 2ªSB"),
                               ("3ªSB", "GRADE 3ªSB"), ("4ªSB", "GRADE 4ªSB"),
                               ("ADM", "GRADE ADM")], start=4):
    ws.cell(i, 1, sb).font = Font(bold=True)
    ws.cell(i, 2).value = f"=IFERROR(VLOOKUP(A{i},PARAMETROS!$A$9:$B$13,2,0),0)"
    ws.cell(i, 3).value = f'=SUM({aba}!$E$3:$E${2 + N}) if False else f"=SUM(\'{aba}\'!$E$3:$E${2 + N})"'
    ws.cell(i, 3).value = f"=SUM('{aba}'!$E$3:$E${2 + N})"
    ws.cell(i, 4).value = f"=B{i}-C{i}"
    ws.cell(i, 5).value = f'=IF(D{i}>0,"SOBRANDO",IF(D{i}<0,"EXCESSO","OK"))'
ws["A10"] = "TOTAL"
ws["A10"].font = Font(bold=True)
ws["B10"] = "=SUM(B4:B8)"
ws["C10"] = "=SUM(C4:C8)"
ws["D10"] = "=B10-C10"
ws["E10"] = '=IF(D10>0,"SOBRANDO",IF(D10<0,"EXCESSO","OK"))'
for col, w in zip("ABCDE", (10, 26, 14, 10, 12)):
    ws.column_dimensions[col].width = w

wb.save(OUT)
print("salvo:", OUT)
print("ADM no mes:", sum(1 for m in mes if m[1] == "ADM"))
