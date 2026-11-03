from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
V2 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v2.xlsx"
V3 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v3.xlsx"
a, b = load_workbook(V2), load_workbook(V3)
res = b["RESUMO"]
nome_para_P = {}
for r in range(2,112):
    p,m,g = (res.cell(r,c).value for c in (2,3,4))
    if all(isinstance(x,str) for x in (p,m,g)):
        nome_para_P[f"{p} {m} {g}"] = res.cell(r,16).value
# confere neutralidade: celulas reescritas cuja formula antiga tinha *2 devem ter P preenchido
diverg = []
for sh,col in [("1ªSB","AL"),("2ªSB","AK"),("3ªSB","AK"),("4ªSB","AK")]:
    ws=b[sh]; wa=a[sh]
    for row in ws.iter_rows():
        for c in row:
            if c.column_letter != col: continue
            novo = c.value.text if isinstance(c.value,ArrayFormula) else c.value
            if not (isinstance(novo,str) and novo.startswith("=IF(XLOOKUP")): continue
            antigo = wa[c.coordinate].value
            antigo = antigo.text if isinstance(antigo,ArrayFormula) else antigo
            nome = ws.cell(c.row,4).value
            temP = bool(nome_para_P.get(nome))
            if "*2" in str(antigo) and not temP:
                diverg.append((sh, c.coordinate, "antigo saldo, P vazio"))
print("divergencias de regime:", diverg if diverg else "nenhuma")
ws = b["4ªSB"]
marcas = {ws.cell(3,c).value for c in range(6,37) if isinstance(ws.cell(3,c).value,str)}
print("marcas do ERALDIR (linha 3):", sorted(marcas))
print("formula dele:", (ws["AK3"].value.text if isinstance(ws["AK3"].value,ArrayFormula) else ws["AK3"].value))
print("P dele:", res.cell(9,16).value)
