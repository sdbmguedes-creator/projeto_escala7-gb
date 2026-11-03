"""Fase 1 da reciprocidade: aba INTEGRACAO com painel por militar."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import FormulaRule

P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_2.0_prototipo.xlsx"
wb = load_workbook(P)

hdr = Font(bold=True, color="FFFFFF")
fill = PatternFill("solid", fgColor="4A6FA5")

OP_N, ADM_N, VIS_N = 32, 14, 12
OP_S, ADM_S, VIS_S = 3, 36, 52

ws = wb.create_sheet("INTEGRACAO")
ws["A1"] = "INTEGRAÇÃO — onde cada militar está escalado"
ws["A1"].font = Font(bold=True, size=14)
ws["B1"] = '=TEXT(PARAMETROS!$B$1,"mmmm/yyyy")'

cab = ["MILITAR", "SB", "DIAS NA SB", "VISITANTE EM", "NO ACO", "ALERTA"]
for i, v in enumerate(cab, start=1):
    c = ws.cell(3, i, v)
    c.font = hdr
    c.fill = fill
    c.alignment = Alignment(horizontal="center")

mes = wb["MES"]
n = mes.max_row - 1
for i in range(n):
    r = 4 + i
    mr = 2 + i
    ws[f"A{r}"] = f"=IF(MES!$D{mr}=\"\",\"\",PESSOAS!$E$2:$E$200)"
    ws[f"A{r}"] = f'=IF(MES!$A{mr}="","",IFERROR(INDEX(PESSOAS!$E$2:$E$200,MATCH(MES!$A{mr},PESSOAS!$A$2:$A$200,0)),""))'
    ws[f"B{r}"] = f'=IF(MES!$D{mr}="","",MES!$D{mr})'
    # dias marcados na propria SB (coluna E das grades = PJES; usamos contagem de marcacoes)
    ws[f"C{r}"] = (
        f'=IF($B{r}="","",IFERROR(IF($B{r}="1ªSB",COUNTA(INDEX(\'GRADE 1ªSB\'!$G${OP_S}:$AK${OP_S+OP_N-1},'
        f'MATCH($A{r},\'GRADE 1ªSB\'!$D${OP_S}:$D${OP_S+OP_N-1},0),0)),'
        f'IF($B{r}="2ªSB",COUNTA(INDEX(\'GRADE 2ªSB\'!$G${OP_S}:$AK${OP_S+OP_N-1},'
        f'MATCH($A{r},\'GRADE 2ªSB\'!$D${OP_S}:$D${OP_S+OP_N-1},0),0)),'
        f'IF($B{r}="3ªSB",COUNTA(INDEX(\'GRADE 3ªSB\'!$G${OP_S}:$AK${OP_S+OP_N-1},'
        f'MATCH($A{r},\'GRADE 3ªSB\'!$D${OP_S}:$D${OP_S+OP_N-1},0),0)),'
        f'IF($B{r}="4ªSB",COUNTA(INDEX(\'GRADE 4ªSB\'!$G${OP_S}:$AK${OP_S+OP_N-1},'
        f'MATCH($A{r},\'GRADE 4ªSB\'!$D${OP_S}:$D${OP_S+OP_N-1},0),0)),0)))),0))'
    )
    # em quantas SBs ele aparece como visitante
    ws[f"D{r}"] = (
        f'=IF($A{r}="","",IF(COUNTIF(\'GRADE 1ªSB\'!$D${VIS_S}:$D${VIS_S+VIS_N-1},$A{r})'
        f'+COUNTIF(\'GRADE 2ªSB\'!$D${VIS_S}:$D${VIS_S+VIS_N-1},$A{r})'
        f'+COUNTIF(\'GRADE 3ªSB\'!$D${VIS_S}:$D${VIS_S+VIS_N-1},$A{r})'
        f'+COUNTIF(\'GRADE 4ªSB\'!$D${VIS_S}:$D${VIS_S+VIS_N-1},$A{r})>0,"SIM",""))'
    )
    ws[f"E{r}"] = f'=IF($A{r}="","",IF(COUNTIF(ACO!$D$7:$D$54,$A{r})>0,"SIM",""))'
    ws[f"F{r}"] = f'=IF(AND($C{r}<>"",$C{r}>0,OR($D{r}="SIM",$E{r}="SIM")),"VERIFICAR","")'

ult = 3 + n
ws.conditional_formatting.add(f"A4:F{ult}", FormulaRule(
    formula=['$F4="VERIFICAR"'], fill=PatternFill("solid", fgColor="FFF4CCCC")))

for col, w in zip("ABCDEF", (34, 10, 12, 14, 10, 12)):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A4"

wb.save(P)
print("INTEGRACAO criada | linhas:", n, "| abas:", len(wb.sheetnames))
