"""Terceira passada: formulas que contam apenas PJ (sem nenhum termo de 12h)."""
import re
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

F = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v2.xlsx"
wb = load_workbook(F)


def t(v):
    return v.text if isinstance(v, ArrayFormula) else v


def put(cell, novo):
    v = cell.value
    cell.value = ArrayFormula(ref=v.ref, text=novo) if isinstance(v, ArrayFormula) else novo


pat = re.compile(r'^(=COUNTIF\((\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+),"PJ"\)\*2)$')
feitos = []
for sh, col in {"1ªSB": "AM", "2ªSB": "AL", "3ªSB": "AL", "4ªSB": "AL"}.items():
    ws = wb[sh]
    for r in range(1, ws.max_row + 1):
        c = ws[f"{col}{r}"]
        x = t(c)
        if not isinstance(x, str):
            continue
        m = pat.match(x.strip())
        if not m:
            continue
        put(c, f'{m.group(1)}+COUNTIF({m.group(2)},"PJ12")+COUNTIF({m.group(2)},"PJ12H")')
        feitos.append(f"{sh}!{c.coordinate}")

wb.save(F)
print("completados agora:", len(feitos))
print(feitos[:5])

# recontagem
wb2 = load_workbook(F)
ok = inc = 0
for sh, col in {"1ªSB": "AM", "2ªSB": "AL", "3ªSB": "AL", "4ªSB": "AL"}.items():
    ws = wb2[sh]
    for r in range(1, ws.max_row + 1):
        x = t(ws[f"{col}{r}"].value)
        if isinstance(x, str) and x.startswith("=COUNTIF") and "*2" in x and '"PJ"' in x:
            if '"PJ12"' in x and '"PJ12H"' in x:
                ok += 1
            else:
                inc += 1
print("PJES completos:", ok, "| incompletos:", inc)
