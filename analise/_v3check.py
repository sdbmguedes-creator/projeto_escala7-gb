from collections import Counter
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
V2 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v2.xlsx"
V3 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v3.xlsx"
a, b = load_workbook(V2), load_workbook(V3)
def t(v): return v.text if isinstance(v,ArrayFormula) else v
res=b["RESUMO"]
print("coluna P:", dict(Counter(str(res.cell(r,16).value) for r in range(2,112) if res.cell(r,16).value is not None)))
nomeP={}
for r in range(2,112):
    p,m,g=(res.cell(r,c).value for c in (2,3,4))
    if all(isinstance(x,str) for x in (p,m,g)): nomeP[f"{p} {m} {g}"]=res.cell(r,16).value
erros=[]
for sh,ci in [("1ªSB",38),("2ªSB",37),("3ªSB",37),("4ªSB",37)]:
    for r in range(1,b[sh].max_row+1):
        novo=t(b[sh].cell(r,ci).value); velho=t(a[sh].cell(r,ci).value)
        if not (isinstance(novo,str) and novo.startswith("=IF(XLOOKUP")): continue
        P=bool(nomeP.get(b[sh].cell(r,4).value))
        saldo_velho = "*2" in str(velho)
        if P != saldo_velho: erros.append((sh,b[sh].cell(r,ci).coordinate,P,saldo_velho))
print("trocas indevidas de regime:", erros if erros else "nenhuma")
print("total alterado vs v2:", sum(1 for sh in b.sheetnames for row in b[sh].iter_rows() for c in row if t(a[sh][c.coordinate].value)!=t(c.value)))
