"""Prototipo 2.0 v7 - ADM dentro da 1aSB, dropdowns e MES com posto/nome."""
import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
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

# layout uniforme
OP_S, OP_N = 3, 32          # operacionais: 3..34
ADM_HDR = 35
ADM_S, ADM_N = 36, 14       # ADM: 36..49
WK = 50                     # auxiliar fim de semana
VIS_TIT = 51
VIS_S, VIS_N = 52, 12       # visitantes: 52..63
R_BM, R_MOT, R_DEF = 65, 66, 67

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
for i, sb in enumerate(["1ªSB", "2ªSB", "3ªSB", "4ªSB", "ADM"]):
    ws.cell(9 + i, 1, sb)
ws["A14"] = "TOTAL"
ws["B14"] = "=SUM(B9:B13)"
ws["A16"] = "Ordem das escalas"
ws["A16"].font = Font(bold=True)
for i, (esc, ordem_esc) in enumerate([("ALFA", 1), ("BRAVO", 2), ("CHARLIE", 3), ("DELTA", 4),
                                     ("ADM", 5), ("X - A", 6), ("X - B", 6), ("X - C", 6),
                                     ("X - D", 6), ("X - ADM", 6)]):
    ws.cell(17 + i, 1, esc)
    ws.cell(17 + i, 2, ordem_esc)
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
ws.append(["ID", "POSTO", "NOME DE GUERRA", "SEÇÃO", "CÓDIGO", "ESCALA",
           "ORDEM ESCALA", "ADM", "MOTORISTA", "ORDEM", "CHAVE ORD", "POSIÇÃO", "CHAVE"])
for c in ws[1]:
    c.font = hdr
    c.fill = fill
for i, m in enumerate(mes, start=2):
    ws.cell(i, 1, m[0])
    ws.cell(i, 2).value = f'=IFERROR(INDEX(PESSOAS!$B$2:$B$200,MATCH($A{i},PESSOAS!$A$2:$A$200,0)),"")'
    ws.cell(i, 3).value = f'=IFERROR(INDEX(PESSOAS!$C$2:$C$200,MATCH($A{i},PESSOAS!$A$2:$A$200,0)),"")'
    ws.cell(i, 4, m[1])
    ws.cell(i, 5).value = f'=IF($D{i}="","",IFERROR(VLOOKUP($D{i},PARAMETROS!$A$4:$D$7,4,0),0))'
    ws.cell(i, 6, m[2])
    ws.cell(i, 7).value = f'=IF($F{i}="","",IFERROR(VLOOKUP($F{i},PARAMETROS!$A$17:$B$26,2,0),9))'
    ws.cell(i, 8, m[3])
    ws.cell(i, 9, m[4])
    ws.cell(i, 10, m[5])
    ws.cell(i, 11).value = f'=IF($G{i}="","",$G{i}*100000+$J{i})'
    ws.cell(i, 12).value = f'=IF($D{i}="","",COUNTIFS($D$2:$D$200,$D{i},$K$2:$K$200,"<"&$K{i})+1)'
    ws.cell(i, 13).value = f'=IF($D{i}="","",$E{i}&"#"&$L{i})'
ws.freeze_panes = "A2"
for col, w in zip("ABCDEFGHIJKLM", (14, 10, 18, 10, 9, 12, 12, 7, 12, 8, 14, 10, 10)):
    ws.column_dimensions[col].width = w

dv_sec = DataValidation(type="list", formula1='"1ªSB,2ªSB,3ªSB,4ªSB,ADM"', allow_blank=True)
ws.add_data_validation(dv_sec)
dv_sec.add("D2:D200")
dv_esc = DataValidation(
    type="list",
    formula1='"ALFA,BRAVO,CHARLIE,DELTA,ADM,X - A,X - B,X - C,X - D,X - ADM"',
    allow_blank=True)
ws.add_data_validation(dv_esc)
dv_esc.add("F2:F200")

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
    (datetime.datetime(2026, 6, 24), "Sao Joao (PE)"), (datetime.datetime(2026, 6, 29), "Sao Pedro (PE)"),
    (datetime.datetime(2026, 9, 7), "Independencia"), (datetime.datetime(2026, 10, 12), "N. S. Aparecida"),
    (datetime.datetime(2026, 11, 2), "Finados"), (datetime.datetime(2026, 11, 15), "Proclamacao"),
    (datetime.datetime(2026, 11, 20), "Consciencia Negra"), (datetime.datetime(2026, 12, 25), "Natal"),
]:
    ws.append([d, n])
    ws.cell(ws.max_row, 1).number_format = "dd/mm/yyyy"
ws.column_dimensions["A"].width = 13
ws.column_dimensions["B"].width = 32


def montar_grade(nome_exib, titulo, codigo, meta_row, com_adm):
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

    def bloco(linha_ini, n, cod, rotulo, com_titulo):
        if com_titulo:
            c = ws.cell(linha_ini - 1, 4, rotulo)
            c.font = Font(bold=True, color="4A6FA5")
        for i in range(n):
            r = linha_ini + i
            ws[f"AN{r}"] = f'=IFERROR(MATCH({cod}&"#{i + 1}",MES!$M$2:$M$200,0),"")'
            ws[f"A{r}"] = f'=IF($AN{r}="","",{i + 1})'
            ws[f"D{r}"] = (f'=IF($AN{r}="","",IFERROR(INDEX(PESSOAS!$E$2:$E$200,'
                           f'MATCH(INDEX(MES!$A$2:$A$200,$AN{r}),PESSOAS!$A$2:$A$200,0)),""))')
            ws[f"B{r}"] = f'=IF($AN{r}="","",INDEX(MES!$F$2:$F$200,$AN{r}))'
            ws[f"C{r}"] = f'=IF($AN{r}="","",INDEX(MES!$I$2:$I$200,$AN{r}))'
            ws[f"E{r}"] = (f'=IF($AN{r}="","",COUNTIF(G{r}:AK{r},"PJ")*2+COUNTIF(G{r}:AK{r},"PJ12")'
                           f'+COUNTIF(G{r}:AK{r},"PJ12H")+COUNTIF(G{r}:AK{r},"SPJ"))')

    bloco(OP_S, OP_N, codigo, "", False)
    if com_adm:
        bloco(ADM_S, ADM_N, 0, "ADMINISTRATIVO (ADM)", True)
        # faixa visual separando os blocos
        for cc in range(1, 38):
            cell = ws.cell(ADM_HDR, cc)
            cell.fill = PatternFill("solid", fgColor="D9E2F3")

    # visitantes: nome escolhido no menu (coluna D), demais colunas por busca
    ws.cell(VIS_TIT, 4, "VISITANTES (militares de outras SBs)").font = Font(bold=True, color="8B4513")
    for cc in range(1, 38):
        ws.cell(VIS_TIT, cc).fill = PatternFill("solid", fgColor="FFF2CC")
    for i in range(VIS_N):
        r = VIS_S + i
        ws[f"D{r}"] = None
        ws[f"B{r}"] = (f'=IF($D{r}="","",IFERROR(INDEX(MES!$F$2:$F$200,'
                       f'MATCH(INDEX(PESSOAS!$A$2:$A$200,MATCH($D{r},PESSOAS!$E$2:$E$200,0)),'
                       f'MES!$A$2:$A$200,0)),""))')
        ws[f"C{r}"] = (f'=IF($D{r}="","",IFERROR(INDEX(MES!$I$2:$I$200,'
                       f'MATCH(INDEX(PESSOAS!$A$2:$A$200,MATCH($D{r},PESSOAS!$E$2:$E$200,0)),'
                       f'MES!$A$2:$A$200,0)),""))')
        ws[f"E{r}"] = (f'=IF($D{r}="","",COUNTIF(G{r}:AK{r},"PJ")*2+COUNTIF(G{r}:AK{r},"PJ12")'
                       f'+COUNTIF(G{r}:AK{r},"PJ12H")+COUNTIF(G{r}:AK{r},"SPJ"))')

    ws[f"D{R_BM}"] = "MILITARES POR DIA"
    ws[f"D{R_MOT}"] = "MOTORISTAS POR DIA"
    ws[f"D{R_DEF}"] = "DÉFICIT / SOBRA"
    for r in (R_BM, R_MOT, R_DEF):
        ws[f"D{r}"].font = Font(bold=True)

    def cont(col, r1, r2):
        b = f"{col}${r1}:{col}${r2}"
        return (f'COUNTIF({b},"x")+COUNTIF({b},"PJ")+COUNTIF({b},"SPJ")'
                f'+(COUNTIF({b},"PJ12")+COUNTIF({b},"PJ12H")+COUNTIF({b},"S1"))*0.5')

    def cont_mot(col, r1, r2):
        b = f"{col}${r1}:{col}${r2}"
        cb = f"$C${r1}:$C${r2}"
        return (f'COUNTIFS({cb},"M",{b},"x")+COUNTIFS({cb},"M",{b},"PJ")+COUNTIFS({cb},"M",{b},"SPJ")'
                f'+(COUNTIFS({cb},"M",{b},"PJ12")+COUNTIFS({cb},"M",{b},"PJ12H")+COUNTIFS({cb},"M",{b},"S1"))*0.5')

    for i in range(31):
        col = get_column_letter(7 + i)
        expr = cont(col, OP_S, OP_S + OP_N - 1) + "+" + cont(col, ADM_S, ADM_S + ADM_N - 1) \
               + "+" + cont(col, VIS_S, VIS_S + VIS_N - 1)
        expr_mot = cont_mot(col, OP_S, OP_S + OP_N - 1) + "+" + cont_mot(col, ADM_S, ADM_S + ADM_N - 1) \
                   + "+" + cont_mot(col, VIS_S, VIS_S + VIS_N - 1)
        ws[f"{col}{R_BM}"] = f'=IF({col}$2="","",{expr})'
        ws[f"{col}{R_MOT}"] = f'=IF({col}$2="","",{expr_mot})'
        if meta_row:
            ws[f"{col}{R_DEF}"] = f'=IF({col}$2="","",PARAMETROS!$B${meta_row}-{col}{R_BM})'
        else:
            ws[f"{col}{R_DEF}"] = '=""'
        ws[f"{col}{WK}"] = (f'=IF({col}$2="","",IF(OR(WEEKDAY({col}$2,2)>5,'
                            f'COUNTIF(FERIADOS!$A:$A,{col}$2)>0),1,0))')
    ws.cell(WK, 4, "aux")
    ws.row_dimensions[WK].hidden = True

    ws.conditional_formatting.add(f"G3:AK{VIS_S + VIS_N - 1}", FormulaRule(
        formula=[f'G${WK}=1'], fill=PatternFill("solid", fgColor="FFF4CCCC")))
    ws.conditional_formatting.add("G2:AK2", FormulaRule(
        formula=[f'G${WK}=1'], fill=PatternFill("solid", fgColor="FFEA9999")))

    # cores por escala (ALFA verde, BRAVO amarelo, CHARLIE azul, DELTA cinza)
    cores = [("ALFA", "FFC6E0B4"), ("BRAVO", "FFFFE699"),
             ("CHARLIE", "FF9FC5E8"), ("DELTA", "FFD9D9D9")]
    for esc, cor in cores:
        ws.conditional_formatting.add(f"A3:AK{VIS_S + VIS_N - 1}", FormulaRule(
            formula=[f'AND($B3="{esc}",$D3<>"")'],
            fill=PatternFill("solid", fgColor=cor)))
    # faixa de inicio de grupo (borda superior simulada por preenchimento mais escuro)
    ws.conditional_formatting.add(f"A3:AK{VIS_S + VIS_N - 1}", FormulaRule(
        formula=['AND($B3<>"",$B3<>$B2)'],
        fill=PatternFill("solid", fgColor="FFBDD7EE")))

    # dropdown dos codigos de servico (evita erro de digitacao)
    dv_cod = DataValidation(type="list", formula1='"X,PJ,PJ12,PJ12H,S1,SPJ,SB,N"',
                            allow_blank=True, showErrorMessage=True, errorStyle="stop",
                            errorTitle="Código inválido",
                            error="Use a lista: X, PJ, PJ12, PJ12H, S1, SPJ, SB ou N.")
    ws.add_data_validation(dv_cod)
    dv_cod.add(f"G3:AK{VIS_S + VIS_N - 1}")

    # dropdown do nome do visitante (lista do cadastro)
    dv_vis = DataValidation(type="list", formula1="PESSOAS!$E$2:$E$200",
                            allow_blank=True, showErrorMessage=True, errorStyle="stop",
                            errorTitle="Selecione um militar do cadastro",
                            error="O nome precisa existir na aba PESSOAS.")
    ws.add_data_validation(dv_vis)
    dv_vis.add(f"D{VIS_S}:D{VIS_S + VIS_N - 1}")

    for col, w in zip(["A", "B", "C", "D", "E", "F"], (4, 9, 6, 30, 6, 5)):
        ws.column_dimensions[col].width = w
    ws.column_dimensions["AN"].hidden = True
    ws.freeze_panes = "G3"


montar_grade("GRADE 1ªSB", "1ªSB — 7º GB — CBMPE", 1, 4, True)
montar_grade("GRADE 2ªSB", "2ªSB — 7º GB — CBMPE", 2, 5, False)
montar_grade("GRADE 3ªSB", "3ªSB — 7º GB — CBMPE", 3, 6, False)
montar_grade("GRADE 4ªSB", "4ªSB — 7º GB — CBMPE", 4, 7, False)

ws = wb.create_sheet("DIARIO")
ws["A1"] = "DIÁRIO — 7º GB"
ws["A1"].font = Font(bold=True, size=14)
ws["B1"] = '=TEXT(PARAMETROS!$B$1,"mmmm/yyyy")'
for i in range(31):
    col = get_column_letter(2 + i)
    c = ws[f"{col}2"]
    g = get_column_letter(7 + i)
    c.value = f"=IF('GRADE 1ªSB'!{g}$2=\"\",\"\",'GRADE 1ªSB'!{g}$2)"
    c.number_format = "dd"
    c.font = hdr
    c.fill = fill
    c.alignment = Alignment(horizontal="center")
    ws.column_dimensions[col].width = 4
ws["A2"] = "Dia"
ws["A2"].font = hdr
ws["A2"].fill = fill

for i, (rot, aba) in enumerate([("1ªSB + ADM", "GRADE 1ªSB"), ("2ªSB", "GRADE 2ªSB"),
                                ("3ªSB", "GRADE 3ªSB"), ("4ªSB", "GRADE 4ªSB")]):
    r = 3 + i
    ws.cell(r, 1, rot).font = Font(bold=True)
    for j in range(31):
        col = get_column_letter(2 + j)
        g = get_column_letter(7 + j)
        ws[f"{col}{r}"] = f"=IF({col}$2=\"\",\"\",'{aba}'!{g}{R_BM})"
ws.cell(7, 1, "TOTAL BM").font = Font(bold=True)
for j in range(31):
    col = get_column_letter(2 + j)
    ws[f"{col}7"] = f"=IF({col}$2=\"\",\"\",SUM({col}3:{col}6))"
ws.cell(8, 1, "MOTORISTAS").font = Font(bold=True)
for j in range(31):
    col = get_column_letter(2 + j)
    g = get_column_letter(7 + j)
    ws[f"{col}8"] = (f"=IF({col}$2=\"\",\"\",'GRADE 1ªSB'!{g}{R_MOT}+'GRADE 2ªSB'!{g}{R_MOT}"
                     f"+'GRADE 3ªSB'!{g}{R_MOT}+'GRADE 4ªSB'!{g}{R_MOT})")
ws.cell(9, 1, "META 7ºGB").font = Font(bold=True)
for j in range(31):
    col = get_column_letter(2 + j)
    ws[f"{col}9"] = "=SUM(PARAMETROS!$B$4:$B$7)"
ws.cell(10, 1, "DÉFICIT").font = Font(bold=True)
for j in range(31):
    col = get_column_letter(2 + j)
    ws[f"{col}10"] = f"=IF({col}$2=\"\",\"\",{col}9-{col}7)"
ws.column_dimensions["A"].width = 16
ws.freeze_panes = "B3"

ws = wb.create_sheet("PJES")
ws["A1"] = "PJES — cotas do mês"
ws["A1"].font = Font(bold=True, size=14)
for i, v in enumerate(["SB", "Cotas recebidas (SEI)", "Utilizadas", "Saldo", "Situação"], start=1):
    c = ws.cell(3, i, v)
    c.font = hdr
    c.fill = fill
for i, sb in enumerate(["1ªSB", "2ªSB", "3ªSB", "4ªSB", "ADM"], start=4):
    ws.cell(i, 1, sb).font = Font(bold=True)
    ws.cell(i, 2).value = f"=IFERROR(VLOOKUP(A{i},PARAMETROS!$A$9:$B$13,2,0),0)"
    aba = "GRADE 1ªSB" if sb in ("1ªSB", "ADM") else f"GRADE {sb}"
    faixa = f"$E$3:$E${OP_S + OP_N - 1}"
    if sb == "ADM":
        ws.cell(i, 3).value = f"=SUM('GRADE 1ªSB'!$E${ADM_S}:$E${ADM_S + ADM_N - 1})"
    else:
        ws.cell(i, 3).value = (f"=SUM('{aba}'!{faixa})"
                               f"+SUM('{aba}'!$E${VIS_S}:$E${VIS_S + VIS_N - 1})")
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
print("ADM:", sum(1 for m in mes if m[1] == "ADM"), "| 1aSB:", sum(1 for m in mes if m[1] == "1ªSB"))
