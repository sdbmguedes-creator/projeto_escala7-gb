import re
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
A = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_BLOCO_A.xlsx"
B = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v2.xlsx"
wa, wb = load_workbook(A), load_workbook(B)
def t(v): return v.text if isinstance(v,ArrayFormula) else v
changed = {}
for sh in wb.sheetnames:
    d=[]
    for row in wb[sh].iter_rows():
        for c in row:
            if t(wa[sh][c.coordinate].value) != t(c.value): d.append(c.coordinate)
    if d: changed[sh]=d
print("celulas alteradas vs Bloco A:", {k:len(v) for k,v in changed.items()}, "| total:", sum(len(v) for v in changed.values()))
# checagens
falta_diario = 0; falta_pjes = 0; ok_pjes = 0
for sh,col in [("1ªSB","AM"),("2ªSB","AL"),("3ªSB","AL"),("4ªSB","AL")]:
    ws=wb[sh]
    for r in range(1, ws.max_row+1):
        x = t(ws[f"{col}{r}"].value)
        if not (isinstance(x,str) and x.startswith("=COUNTIF") and "*2" in x and '"PJ"' in x): continue
        tem12 = '"PJ12"' in x; temh = '"PJ12H"' in x
        if tem12 and temh: ok_pjes+=1
        else: falta_pjes+=1
for sh in ["1ªSB","2ªSB","3ªSB","4ªSB"]:
    for row in wb[sh].iter_rows():
        for c in row:
            x=t(c.value)
            if isinstance(x,str) and '"PJ12";"S1"' in x: falta_diario+=1
print("PJES com os dois codigos:", ok_pjes, "| ainda incompletos:", falta_pjes)
print("formulas de meta diaria ainda sem PJ12H:", falta_diario)
# numeracao dos visitantes
ws=wb["1ªSB"]
print("A55:", t(ws["A55"].value))
print("A55:A92 corrigidas:", sum(1 for r in range(55,93) if isinstance(t(ws[f'A{r}'].value),str) and 'RESUMO!$A$2:$A$111' in t(ws[f'A{r}'].value)))
print("orfas:", ws["B29"].value, ws["C29"].value, wb["ACO"]["A54"].value, wb["ACO"]["B54"].value)
