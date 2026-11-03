"""Adiciona a aba ACO ao prototipo (oficiais, militares e marcacao 12/24)."""
import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_2.0_prototipo.xlsx"
wb = load_workbook(P)

hdr = Font(bold=True, color="FFFFFF")
fill = PatternFill("solid", fgColor="4A6FA5")
N_ACO = 48

ws = wb.create_sheet("ACO")
ws["A1"] = "ACO — 7º GB — CBMPE"
ws["A1"].font = Font(bold=True, size=14)
ws["B1"] = '=TEXT(PARAMETROS!$B$1,"mmmm/yyyy")'
ws["B1"].font = Font(bold=True, size=12)

# oficiais: linhas 3 e 4 (1o e 2o turno), nomes digitados
ws["A3"] = "1º TURNO"
ws["A3"].font = Font(bold=True)
ws["A4"] = "2º TURNO"
ws["A4"].font = Font(bold=True)
for i in range(31):
    col = get_column_letter(6 + i)
    for r in (3, 4):
        c = ws[f"{col}{r}"]
        c.font = hdr
        c.fill = PatternFill("solid", fgColor="8B4513")
        c.alignment = Alignment(horizontal="center")
        ws.column_dimensions[col].width = 11

# cabecalho dos militares
cab = {1: "SB", 2: "ESCALA", 3: "F", 4: "MILITAR", 5: "VOLUNT."}
for col, v in cab.items():
    c = ws.cell(6, col, v)
    c.font = hdr
    c.fill = fill
    c.alignment = Alignment(horizontal="center")
for i in range(31):
    col = get_column_letter(6 + i)
    c = ws[f"{col}6"]
    c.value = (f'=IF(MONTH(PARAMETROS!$B$1+{i})=MONTH(PARAMETROS!$B$1),'
               f'PARAMETROS!$B$1+{i},"")')
    c.number_format = "dd"
    c.font = hdr
    c.fill = fill
    c.alignment = Alignment(horizontal="center")

# linhas de militares do ACO (nome escolhido no menu)
for i in range(N_ACO):
    r = 7 + i
    ws[f"A{r}"] = (f'=IF($D{r}="","",IFERROR(INDEX(MES!$D$2:$D$200,'
                   f'MATCH(INDEX(PESSOAS!$A$2:$A$200,MATCH($D{r},PESSOAS!$E$2:$E$200,0)),'
                   f'MES!$A$2:$A$200,0)),""))')
    ws[f"B{r}"] = (f'=IF($D{r}="","",IFERROR(INDEX(MES!$F$2:$F$200,'
                   f'MATCH(INDEX(PESSOAS!$A$2:$A$200,MATCH($D{r},PESSOAS!$E$2:$E$200,0)),'
                   f'MES!$A$2:$A$200,0)),""))')

ws.column_dimensions["A"].width = 8
ws.column_dimensions["B"].width = 10
ws.column_dimensions["C"].width = 4
ws.column_dimensions["D"].width = 32
ws.column_dimensions["E"].width = 9

r_tot = 7 + N_ACO + 2
ws.cell(r_tot, 4, "TOTAL 12h").font = Font(bold=True)
ws.cell(r_tot + 1, 4, "TOTAL 24h").font = Font(bold=True)
ws.cell(r_tot, 4).font = Font(bold=True)
for i in range(31):
    col = get_column_letter(6 + i)
    base = f"{col}$7:{col}${6 + N_ACO}"
    ws[f"{col}{r_tot}"] = f'=IF({col}$6="","",COUNTIF({base},12*1))'
    ws[f"{col}{r_tot + 1}"] = f'=IF({col}$6="","",COUNTIF({base},24*1))'

dv_cod = DataValidation(type="list", formula1='"12,24"', allow_blank=True,
                        showErrorMessage=True, errorStyle="stop",
                        errorTitle="Turno inválido", error="Use 12 ou 24.")
ws.add_data_validation(dv_cod)
dv_cod.add(f"F7:AJ{6 + N_ACO}")

dv_nome = DataValidation(type="list", formula1="PESSOAS!$E$2:$E$200", allow_blank=True,
                         showErrorMessage=True, errorStyle="stop",
                         errorTitle="Selecione um militar do cadastro",
                         error="O nome precisa existir na aba PESSOAS.")
ws.add_data_validation(dv_nome)
dv_nome.add(f"D7:D{6 + N_ACO}")

ws.freeze_panes = "F7"
wb.save(P)
print("aba ACO criada | abas:", len(wb.sheetnames))
