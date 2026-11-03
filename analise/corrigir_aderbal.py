"""Corrige marcacoes do ADERBAL e os #N/A, gerando o arquivo pronto."""
import re
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

SRC = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11_V2.xlsx"
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_corrigida_01-11.xlsx"

wb = load_workbook(SRC)
log = []


def t(v):
    return v.text if isinstance(v, ArrayFormula) else v


def put(cell, novo):
    v = cell.value
    cell.value = ArrayFormula(ref=v.ref, text=novo) if isinstance(v, ArrayFormula) else novo


# ---- 1) 4aSB linha 3: limpar marcacoes e relancar as corretas ----
ws4 = wb["4ªSB"]
for c in range(6, 37):          # F..AJ = dias 1..31
    if ws4.cell(3, c).value is not None:
        ws4.cell(3, c).value = None
        log.append(f"4ªSB!{ws4.cell(3, c).coordinate} limpo")

MARKS = {"I3": "X", "M3": "X", "O3": "PJ", "Q3": "X",
         "U3": "X", "Y3": "X", "AC3": "X", "AG3": "X"}
for cell, cod in MARKS.items():
    ws4[cell].value = cod
    log.append(f"4ªSB!{cell} = {cod}")

# ---- 2) 1aSB: B3/C3 com IFERROR (vaga vazia) ----
ws1 = wb["1ªSB"]
for cell, idx in [("B3", 3), ("C3", 2)]:
    atual = t(ws1[cell].value)
    if isinstance(atual, str) and atual.startswith("=VLOOKUP"):
        put(ws1[cell], f'=IFERROR(VLOOKUP($D3,RESUMO!$F$2:$H$400,{idx},0),"")')
        log.append(f"1ªSB!{cell} protegido com IFERROR")

# ---- 3) 4aSB: B13/C13 da linha vazia ----
for cell in ["B13", "C13"]:
    if ws4[cell].value is not None:
        ws4[cell].value = None
        log.append(f"4ªSB!{cell} limpo")

wb.save(OUT)
print("alteracoes:", len(log))
for x in log:
    print("  ", x)
