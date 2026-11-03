"""Normaliza espacos em nomes, de forma consistente entre as abas."""
import re
from openpyxl import load_workbook

SRC = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_conferencias_01-11.xlsx"
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_final_01-11.xlsx"

PADRAO = re.compile(r"\d{6}-\d")
wb = load_workbook(SRC)
alterados = []


def norm(s):
    return " ".join(s.split())


# 1) RESUMO: posto (B), matricula (C) e nome de guerra (D)
res = wb["RESUMO"]
for r in range(2, 115):
    for col in (2, 3, 4):
        c = res.cell(r, col)
        if isinstance(c.value, str):
            novo = norm(c.value)
            if novo != c.value:
                alterados.append(f"RESUMO!{c.coordinate}: {c.value!r} -> {novo!r}")
                c.value = novo

# 2) abas de SB: coluna D
for sh in ["1ªSB", "2ªSB", "3ªSB", "4ªSB"]:
    ws = wb[sh]
    for r in range(3, 130):
        c = ws.cell(r, 4)
        if isinstance(c.value, str) and PADRAO.search(c.value):
            novo = norm(c.value)
            if novo != c.value:
                alterados.append(f"{sh}!{c.coordinate}: {c.value!r} -> {novo!r}")
                c.value = novo

# 3) ACO
ws = wb["ACO"]
for r in range(6, 60):
    c = ws.cell(r, 4)
    if isinstance(c.value, str) and PADRAO.search(c.value):
        novo = norm(c.value)
        if novo != c.value:
            alterados.append(f"ACO!{c.coordinate}: {c.value!r} -> {novo!r}")
            c.value = novo

# 4) Planilha de PJES (coluna C = nome) e PJES DISTRIBUICAO
ws = wb["Planilha de PJES"]
for r in range(2, 260):
    c = ws.cell(r, 3)
    if isinstance(c.value, str) and PADRAO.search(c.value):
        novo = norm(c.value)
        if novo != c.value:
            alterados.append(f"Planilha de PJES!{c.coordinate}: {c.value!r} -> {novo!r}")
            c.value = novo

ws = wb["PJES DISTRIBUIÇÃO"]
for row in ws.iter_rows():
    for c in row:
        if isinstance(c.value, str) and PADRAO.search(c.value):
            novo = norm(c.value)
            if novo != c.value:
                alterados.append(f"PJES DISTRIBUIÇÃO!{c.coordinate}: {c.value!r} -> {novo!r}")
                c.value = novo

wb.save(OUT)
print("celulas normalizadas:", len(alterados))
for a in alterados:
    print("  ", a)
