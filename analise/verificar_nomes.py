"""Verifica se todos os nomes das abas ainda casam com o cadastro."""
import re
from openpyxl import load_workbook

F = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_final_01-11.xlsx"
wb = load_workbook(F)
PADRAO = re.compile(r"\d{6}-\d")

res = wb["RESUMO"]
cadastro = set()
for r in range(2, 112):
    b, c, d = (res.cell(r, col).value for col in (2, 3, 4))
    if all(isinstance(x, str) for x in (b, c, d)):
        cadastro.add(f"{b} {c} {d}")

print("nomes no cadastro (RESUMO):", len(cadastro))

falhas = []
for sh, col, a, b_ in [("1ªSB", 4, 3, 130), ("2ªSB", 4, 3, 130), ("3ªSB", 4, 3, 130),
                       ("4ªSB", 4, 3, 130), ("ACO", 4, 6, 60)]:
    ws = wb[sh]
    for r in range(a, b_ + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and PADRAO.search(v) and v not in cadastro:
            falhas.append(f"{sh}!row{r}: {v!r}")

ws = wb["Planilha de PJES"]
for r in range(2, 260):
    v = ws.cell(r, 3).value
    if isinstance(v, str) and PADRAO.search(v) and v not in cadastro:
        falhas.append(f"Planilha de PJES!C{r}: {v!r}")

print("nomes que NAO casam com o cadastro:", len(falhas))
for f in falhas[:15]:
    print("   ", f)

# checagem de espacos restantes
resto = 0
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and PADRAO.search(c.value):
                if c.value != " ".join(c.value.split()):
                    resto += 1
                    print("   espaco restante:", ws.title, c.coordinate, repr(c.value))
print("celulas de nome ainda com espaco irregular:", resto)
