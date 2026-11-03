import re
from collections import Counter
from openpyxl import load_workbook
P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
wb = load_workbook(P); wbv = load_workbook(P, data_only=True)
print("formulas por aba e refs cruzadas")
tot=Counter()
for ws in wb.worksheets:
    n=0; cross=0; vres=0; lit=0
    for row in ws.iter_rows():
        for c in row:
            v=c.value
            t = v.text if hasattr(v,"text") else v
            if isinstance(t,str) and t.startswith("="):
                n+=1
                if "!" in t: cross+=1
                if "RESUMO!" in t: vres+=1
                if re.search(r"\$?\d+:\$?[A-Z]*\$?\d+", t) and "RESUMO" in t: lit+=1
    tot[ws.title]=(n,cross,vres)
    print(f"  {ws.title:22} formulas={n:5} cross-sheet={cross:5} refs RESUMO={vres:4}")
print("TOTAL formulas:", sum(v[0] for v in tot.values()), "cross:", sum(v[1] for v in tot.values()))
# celulas de marcacao (X/PJ/...) por aba
for sh in ["1ªSB","2ªSB","3ªSB","4ªSB"]:
    ws=wbv[sh]; n=0
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value,str) and c.value.strip() in {"X","PJ","PJ12","PJ12H","S1","SPJ","N","12","24"}: n+=1
    print(f"  celulas de marcacao {sh}: {n}")
