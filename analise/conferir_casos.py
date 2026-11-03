"""Confere dois casos apontados pela simulacao."""
from openpyxl import load_workbook

V2 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11_V2.xlsx"
wv = load_workbook(V2, data_only=True)

print("=== 2ªSB: ABREU, dias 13 a 18 ===")
ws = wv["2ªSB"]
for r in range(3, 33):
    if ws.cell(r, 4).value and "ABREU" in str(ws.cell(r, 4).value):
        linha = r
        break
for r in [linha]:
    print("  militar:", ws.cell(r, 4).value, "| escala:", ws.cell(r, 2).value)
    for ci in range(ws["F1"].column, ws["AJ1"].column + 1):
        d = ws.cell(2, ci).value
        if d is not None and 13 <= d.day <= 18:
            print(f"    dia {d.day:2}: {ws.cell(r, ci).value!r}")

print()
print("=== 4ªSB: CLEITON, dias 1 a 13 ===")
ws4 = wv["4ªSB"]
for r in range(3, 34):
    if ws4.cell(r, 4).value and "CLEITON" in str(ws4.cell(r, 4).value):
        l4 = r
        break
print("  militar:", ws4.cell(l4, 4).value, "| escala:", ws4.cell(l4, 2).value)
for ci in range(ws4["F1"].column, ws4["AJ1"].column + 1):
    d = ws4.cell(2, ci).value
    if d is not None and 1 <= d.day <= 13:
        print(f"    dia {d.day:2}: {ws4.cell(l4, ci).value!r}")
