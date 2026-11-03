import re
from collections import Counter
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

F = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11_V2.xlsx"
wb = load_workbook(F)
wv = load_workbook(F, data_only=True)


def t(v):
    return v.text if isinstance(v, ArrayFormula) else v


res, resv = wb["RESUMO"], wv["RESUMO"]
print("S9 =", resv["S9"].value, "| S4:S8 =", [resv[f"S{r}"].value for r in range(4, 9)])

err = {}
for ws in wv.worksheets:
    n = sum(1 for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and c.value.startswith("#"))
    if n:
        err[ws.title] = n
print("erros:", err)

print()
print("=== duplicidade de nome por aba ===")
for sh in ["1ªSB", "2ªSB", "3ªSB", "4ªSB"]:
    wsv = wv[sh]
    nomes = [(r, wsv.cell(r, 4).value) for r in range(3, 120)
             if isinstance(wsv.cell(r, 4).value, str) and re.search(r"\d{6}-\d", wsv.cell(r, 4).value)]
    cnt = Counter(n for _, n in nomes)
    dups = {n: [r for r, x in nomes if x == n] for n, c in cnt.items() if c > 1}
    print(f"  {sh}: {len(nomes)} | duplicados: {dups if dups else 'nenhum'}")

print()
print("=== 4aSB: linhas 3 a 32 ===")
ws4, wsv4 = wb["4ªSB"], wv["4ªSB"]
for r in range(3, 33):
    d = wsv4.cell(r, 4).value
    al = wsv4.cell(r, 38).value
    f = t(ws4.cell(r, 38).value)
    n_mark = sum(1 for c in range(6, 37) if isinstance(wsv4.cell(r, c).value, str) and wsv4.cell(r, c).value.strip())
    if isinstance(d, str) or al:
        print(f"  {r:3}: {str(d)[:30]:32} marcas={n_mark:3} AL={al!r:6} | {str(f)[:100]}")
