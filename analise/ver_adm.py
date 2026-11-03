from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

F = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11.xlsx"
wb = load_workbook(F)
wv = load_workbook(F, data_only=True)


def t(v):
    return v.text if isinstance(v, ArrayFormula) else v


ws, wsv = wb["1ªSB"], wv["1ªSB"]
print("=== 1aSB linhas 36-56 ===")
for r in range(36, 57):
    d = wsv.cell(r, 4).value
    b = t(ws.cell(r, 2).value)
    al = wsv.cell(r, 38).value
    if d or (isinstance(b, str) and b.startswith("=")):
        print(f" {r:3}: D={str(d)[:34]:36} AL={al!r:6} B={str(b)[:60]}")

print()
print("=== celulas de total de PJES por SB ===")
for sh, cells in [("1ªSB", ["AM32", "AM51", "AN51"]), ("2ªSB", ["AL31"]),
                  ("3ªSB", ["AL19"]), ("4ªSB", ["AL33"]), ("ACO", ["AK57"])]:
    for c in cells:
        v = t(wb[sh][c].value)
        print(f"  {sh}!{c}: valor={wv[sh][c].value!r} | formula={str(v)[:110]}")
