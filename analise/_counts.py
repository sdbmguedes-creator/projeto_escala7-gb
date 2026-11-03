from collections import Counter
from openpyxl import load_workbook
P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
wbv = load_workbook(P, data_only=True)
ws = wbv["RESUMO"]
rows = [r for r in range(2, 112) if ws.cell(r,5).value]
print("RESUMO linhas com E preenchido:", len(rows), "de", 2, "a", 111)
print("por secao:", Counter(ws.cell(r,5).value for r in rows))
print("por posto:", Counter(ws.cell(r,2).value for r in rows))
print("mulheres (G=F):", sum(1 for r in rows if ws.cell(r,7).value=="F"))
print("soma PJES TOTAL (O):", sum(ws.cell(r,15).value or 0 for r in rows))
print("com PJES>0:", sum(1 for r in rows if (ws.cell(r,15).value or 0)>0))
print("escalas:", Counter(ws.cell(r,8).value for r in rows))
print("Ef Pracas X vs Z:")
for r in range(22,29):
    print("  row",r, ws.cell(r,19).value, "X=", ws.cell(r,24).value, "Z=", ws.cell(r,26).value)
print("1aSB totais: AM30=", wbv["1ªSB"]["AM30"].value, "AL30=", wbv["1ªSB"]["AL30"].value,
      "AM32=", wbv["1ªSB"]["AM32"].value, "AM50=", wbv["1ªSB"]["AM50"].value,
      "AM51=", wbv["1ªSB"]["AM51"].value, "AN51=", wbv["1ªSB"]["AN51"].value)
