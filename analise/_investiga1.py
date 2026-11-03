import re
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
V5 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v5.xlsx"
wb = load_workbook(V5)
print("=== celulas ORDINARIO que sao ARRAY FORMULA (possivel causa do #VALUE) ===")
n=0
for sh,colidx in [("1ªSB",38),("2ªSB",37),("3ªSB",37),("4ªSB",37)]:
    ws=wb[sh]
    for r in range(1, ws.max_row+1):
        v=ws.cell(r,colidx).value
        if isinstance(v, ArrayFormula):
            n+=1
            if n<=6: print(f"  {sh}!{ws.cell(r,colidx).coordinate}: ref={v.ref} | {v.text[:70]}")
print("  total de array formulas na coluna ORDINARIO:", n)
print()
print("=== celulas ORDINARIO comuns (nao-array): quantas ===")
m=0
for sh,colidx in [("1ªSB",38),("2ªSB",37),("3ªSB",37),("4ªSB",37)]:
    ws=wb[sh]
    for r in range(1, ws.max_row+1):
        v=ws.cell(r,colidx).value
        if isinstance(v,str) and v.startswith("=IF(XLOOKUP"): m+=1
print("  ", m)
