from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
V2 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v2.xlsx"
V3 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v3.xlsx"
a, b = load_workbook(V2), load_workbook(V3)
def t(v): return v.text if isinstance(v,ArrayFormula) else v
res = b["RESUMO"]
nome_P = {}
for r in range(2,112):
    p,m,g = (res.cell(r,c).value for c in (2,3,4))
    if all(isinstance(x,str) for x in (p,m,g)):
        nome_P[f"{p} {m} {g}"] = res.cell(r,16).value
diverg=[]
for sh, colidx in [("1ªSB",38),("2ªSB",37),("3ªSB",37),("4ªSB",37)]:
    ws, wa = b[sh], a[sh]
    for r in range(1, ws.max_row+1):
        novo, antigo = t(ws.cell(r,colidx).value), t(wa.cell(r,colidx).value)
        if not (isinstance(novo,str) and novo.startswith("=IF(XLOOKUP")): continue
        temP = bool(nome_P.get(ws.cell(r,4).value))
        if "*2" in str(antigo) and not temP:
            diverg.append((sh, ws.cell(r,colidx).coordinate))
print("divergencias (antigo saldo, P vazio):", diverg if diverg else "nenhuma")
ws=b["4ªSB"]
marcas={ws.cell(3,c).value for c in range(6,37) if isinstance(ws.cell(3,c).value,str)}
print("marcas ERALDIR:", sorted(marcas), "| P:", res.cell(9,16).value)
print("formula ERALDIR:", t(ws.cell(3,37).value))
print("P de nao-ADM (linha 3):", res.cell(3,16).value)
