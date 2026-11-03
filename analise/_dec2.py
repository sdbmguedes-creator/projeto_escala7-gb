from openpyxl import load_workbook
V = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v4.xlsx"
wb = load_workbook(V, data_only=True)
res = wb["RESUMO"]
def nome(r):
    p,m,g = (res.cell(r,c).value for c in (2,3,4))
    return f"{p} {m} {g}" if all(isinstance(x,str) for x in (p,m,g)) else None

print("=== (2) RESUMO!I (ORD) vs coluna ORDINARIO da SB ===")
# mapa nome -> valor ORD na aba
sbval = {}
for sh, col, a, b in [("1ªSB","AL",3,92),("2ªSB","AK",3,53),("3ªSB","AK",3,41),("4ªSB","AK",3,55)]:
    ws = wb[sh]
    for r in range(a,b+1):
        n = ws.cell(r,4).value
        if isinstance(n,str): sbval.setdefault(n, []).append((sh, ws[f"{col}{r}"].value))
difs=[]
for r in range(2,112):
    n = nome(r)
    if not n: continue
    i = res.cell(r,9).value
    v = sbval.get(n, [])
    if not v: continue
    val = v[0][1]
    if i != val: difs.append((n, i, val, v[0][0]))
print("pessoas comparadas:", sum(1 for r in range(2,112) if nome(r)))
print("divergencias RESUMO!I x SB:", len(difs))
for d in difs[:12]: print("   ", d)
