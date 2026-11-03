from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

NEW = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11_V2.xlsx"
wb = load_workbook(NEW)
wv = load_workbook(NEW, data_only=True)


def t(v):
    return v.text if isinstance(v, ArrayFormula) else v


def marcas(sh, r, c1, c2):
    wsv = wv[sh]
    return [(wsv.cell(r, c).coordinate, wsv.cell(r, c).value) for c in range(c1, c2 + 1)
            if isinstance(wsv.cell(r, c).value, str) and wsv.cell(r, c).value.strip()]


print("=== 4aSB linha 3 (vaga do ADM / ADERBAL) ===")
for c, v in marcas("4ªSB", 3, 6, 36):
    print(f"   {c} = {v}")
print("  AL3 =", wv["4ªSB"].cell(3, 38).value)

print()
print("=== 4aSB linha 13 (o que restou) ===")
print("  D13 =", repr(wv["4ªSB"].cell(13, 4).value))
print("  marcas:", marcas("4ªSB", 13, 6, 36))
print("  AL13 =", wv["4ªSB"].cell(13, 38).value)

print()
print("=== 1aSB: ADERBAL no bloco de visitantes ===")
ws1 = wv["1ªSB"]
for r in range(50, 100):
    if ws1.cell(r, 4).value == "CB 718087-0 ADERBAL":
        print(f"  linha {r}: marcas=", marcas("1ªSB", r, 7, 37), "| AM=", ws1.cell(r, 39).value)

print()
print("=== os 2 erros de cada aba ===")
for sh in ["1ªSB", "4ªSB"]:
    for row in wv[sh].iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("#"):
                print(f"  {sh}!{c.coordinate} = {c.value} | {str(t(wb[sh][c.coordinate].value))[:110]}")
