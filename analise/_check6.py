import re
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
wb = load_workbook(P)
def text(v): return v.text if isinstance(v,ArrayFormula) else v
print("coluna ORDINARIO: quantas linhas multiplicam X por 2")
for sh,col,a,b in [("1ªSB","AL",3,92),("2ªSB","AK",3,53),("3ªSB","AK",3,41),("4ªSB","AK",3,55)]:
    ws=wb[sh]; com=[]; sem=[]
    for r in range(a,b+1):
        t=text(ws[f"{col}{r}"].value)
        if not isinstance(t,str) or not t.startswith("=COUNTIF"): continue
        (com if "*2" in t else sem).append(r)
    print(f" {sh}!{col}: com *2 = {len(com)} linhas | sem *2 = {len(sem)} linhas | sem*2 nas linhas {sem[:12]}")
    if com: print(f"     exemplo com *2: {sh}!{col}{com[0]} = {text(ws[f'{col}{com[0]}'].value)[:100]}")
    if sem: print(f"     exemplo sem *2: {sh}!{col}{sem[0]} = {text(ws[f'{col}{sem[0]}'].value)[:100]}")
