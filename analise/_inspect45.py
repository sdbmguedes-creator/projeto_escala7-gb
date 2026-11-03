from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
from collections import Counter
B = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v2.xlsx"
wb = load_workbook(B)
def t(v): return v.text if isinstance(v,ArrayFormula) else v
pad = Counter(); ex = {}
for sh,col in [("1ªSB","AM"),("2ªSB","AL"),("3ªSB","AL"),("4ªSB","AL")]:
    ws=wb[sh]
    for r in range(1, ws.max_row+1):
        x = t(ws[f"{col}{r}"].value)
        if not (isinstance(x,str) and x.startswith("=COUNTIF") and "*2" in x and '"PJ"' in x): continue
        if '"PJ12"' in x and '"PJ12H"' in x: continue
        # normaliza numeros de linha para agrupar
        import re
        k = re.sub(r"\d+", "#", x)
        pad[k]+=1; ex.setdefault(k, f"{sh}!{col}{r}")
for k,n in pad.most_common(10):
    print(f"x{n}  {ex[k]}")
    print(f"      {k[:180]}")
