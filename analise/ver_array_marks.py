from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

V2 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11_V2.xlsx"
wb = load_workbook(V2)
ws = wb["1ªSB"]

for c in ["H73", "J73", "P73", "V73", "AB73", "AD73", "AH73"]:
    v = ws[c].value
    if isinstance(v, ArrayFormula):
        print(f"  {c}: ARRAY ref={v.ref} text={v.text!r}")
    else:
        print(f"  {c}: {v!r}")

print()
print("=== quantas ArrayFormula existem nas linhas de visitantes da 1aSB? ===")
n = 0
for r in range(56, 93):
    for ci in range(7, 38):
        if isinstance(ws.cell(r, ci).value, ArrayFormula):
            n += 1
print("  total:", n)
