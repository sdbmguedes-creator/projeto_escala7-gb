import re
from collections import Counter
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

F = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11.xlsx"
wb = load_workbook(F)
wv = load_workbook(F, data_only=True)


def t(v):
    return v.text if isinstance(v, ArrayFormula) else v


print("=== duplicidade: mesmo nome em mais de uma linha da mesma aba ===")
for sh in ["1ªSB", "2ªSB", "3ªSB", "4ªSB"]:
    wsv = wv[sh]
    nomes = [(r, wsv.cell(r, 4).value) for r in range(3, 100)
             if isinstance(wsv.cell(r, 4).value, str) and re.search(r"\d{6}-\d", wsv.cell(r, 4).value)]
    cnt = Counter(n for _, n in nomes)
    dups = {n: [r for r, x in nomes if x == n] for n, c in cnt.items() if c > 1}
    print(f"  {sh}: {len(nomes)} nomes | duplicados: {dups if dups else 'nenhum'}")

print()
print("=== linhas de total: a partir de qual linha somam? ===")
for sh, rows in [("1ªSB", [31, 32]), ("2ªSB", [32, 33]), ("3ªSB", [20, 21]), ("4ªSB", [33, 34])]:
    ws = wb[sh]
    for r in rows:
        f = t(ws.cell(r, 7).value if sh == "1ªSB" else ws.cell(r, 6).value)
        if isinstance(f, str):
            print(f"  {sh} linha {r}: {f[:150]}")
