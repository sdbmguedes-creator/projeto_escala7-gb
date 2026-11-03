from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

V2 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11_V2.xlsx"
wb = load_workbook(V2)
wv = load_workbook(V2, data_only=True)


def t(v):
    return v.text if isinstance(v, ArrayFormula) else v


ws, wsv = wb["1ªSB"], wv["1ªSB"]
print("=== linha 73 (ADERBAL, visitante na 1aSB) ===")
print("  D73:", wsv.cell(73, 4).value)
print("  B73:", wsv.cell(73, 2).value)
print("  AM73 formula:", t(ws.cell(73, 39).value))
print("  AM73 valor  :", wsv.cell(73, 39).value)
print()
print("=== comparando com outros visitantes (linhas 56-75) ===")
for r in range(56, 76):
    nome = wsv.cell(r, 4).value
    if not isinstance(nome, str):
        continue
    cods = [wsv.cell(r, c).value for c in range(7, 38) if isinstance(wsv.cell(r, c).value, str)]
    print(f"  {r}: {nome[:26]:28} AM={wsv.cell(r,39).value!r:5} marcas={cods}")
