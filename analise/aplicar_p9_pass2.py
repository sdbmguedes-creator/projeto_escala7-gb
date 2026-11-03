"""Segunda passada do P9: linhas sem SPJ na coluna de PJES."""
import re
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

F = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v2.xlsx"
wb = load_workbook(F)


def get_text(cell):
    v = cell.value
    return v.text if isinstance(v, ArrayFormula) else v


def put_text(cell, novo):
    v = cell.value
    if isinstance(v, ArrayFormula):
        cell.value = ArrayFormula(ref=v.ref, text=novo)
    else:
        cell.value = novo


alvo = {"1ªSB": "AM", "2ªSB": "AL", "3ªSB": "AL", "4ªSB": "AL"}
pat = re.compile(r'COUNTIF\((\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+),"([A-Z0-9]+)"\)\s*$')
feitos, jaok = [], []
for sh, col in alvo.items():
    ws = wb[sh]
    for r in range(1, ws.max_row + 1):
        c = ws[f"{col}{r}"]
        t = get_text(c)
        if not (isinstance(t, str) and t.startswith("=COUNTIF") and "*2" in t and '"PJ"' in t):
            continue
        m = pat.search(t)
        if not m:
            continue
        rng, ultimo = m.group(1), m.group(2)
        falta = {"PJ12": "PJ12H", "PJ12H": "PJ12"}.get(ultimo)
        if not falta:
            continue
        if f'"{falta}"' in t:
            jaok.append(f"{sh}!{c.coordinate}")
            continue
        put_text(c, t + f'+COUNTIF({rng},"{falta}")')
        feitos.append(f"{sh}!{c.coordinate}")

wb.save(F)
print("acrescentados agora:", len(feitos))
print("ja estavam completos:", len(jaok))
print("exemplos:", feitos[:5])
