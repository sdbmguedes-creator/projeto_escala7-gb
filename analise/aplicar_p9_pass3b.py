"""Passo 3 (corrigido): formulas que contam apenas PJ, sem termo de 12h."""
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

F = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v2.xlsx"
wb = load_workbook(F)


def t(v):
    return v.text if isinstance(v, ArrayFormula) else v


def put(cell, novo):
    v = cell.value
    cell.value = ArrayFormula(ref=v.ref, text=novo) if isinstance(v, ArrayFormula) else novo


feitos = []
for sh, col in {"1ªSB": "AM", "2ªSB": "AL", "3ªSB": "AL", "4ªSB": "AL"}.items():
    ws = wb[sh]
    for r in range(1, ws.max_row + 1):
        c = ws[f"{col}{r}"]
        x = t(c)
        if not isinstance(x, str):
            continue
        s = x.strip()
        if not s.startswith("=COUNTIF("):
            continue
        if not s.endswith('"PJ")*2'):
            continue
        rng = s[len("=COUNTIF("):s.index(",")]
        put(c, s + f'+COUNTIF({rng},"PJ12")+COUNTIF({rng},"PJ12H")')
        feitos.append(f"{sh}!{c.coordinate}")

wb.save(F)
print("completados agora:", len(feitos))
print(feitos[:6])

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
