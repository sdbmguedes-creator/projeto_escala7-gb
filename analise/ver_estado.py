from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

F = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11.xlsx"
wb = load_workbook(F)
wv = load_workbook(F, data_only=True)


def t(v):
    return v.text if isinstance(v, ArrayFormula) else v


print("=== erros em 1aSB e RESUMO ===")
for sh in ["1ªSB", "RESUMO"]:
    cells = [c.coordinate for row in wv[sh].iter_rows() for c in row
             if isinstance(c.value, str) and c.value.startswith("#")]
    print(f" {sh}: {len(cells)} -> {cells[:8]}")
    if cells:
        c0 = cells[0]
        print(f"    valor {c0}={wv[sh][c0].value} | formula: {str(t(wb[sh][c0].value))[:150]}")

print()
print("=== 4aSB: linhas com nome e a coluna AL (PJES) ===")
ws, wsv = wb["4ªSB"], wv["4ªSB"]
for r in range(1, 60):
    d = wsv.cell(r, 4).value
    al = wsv.cell(r, 38).value
    f = t(ws.cell(r, 38).value)
    if isinstance(d, str) and d.strip():
        print(f"  {r:3}: {str(d)[:30]:32} AL={al} | {str(f)[:90]}")
